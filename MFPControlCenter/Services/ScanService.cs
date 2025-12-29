using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using MFPControlCenter.Models;

namespace MFPControlCenter.Services
{
    public class ScanService
    {
        // WIA property IDs
        private const string WIA_IPA_DATATYPE = "6146";
        private const string WIA_IPS_XRES = "6147";
        private const string WIA_IPS_YRES = "6148";
        private const string WIA_IPS_XEXTENT = "6151";
        private const string WIA_IPS_YEXTENT = "6152";
        private const string WIA_IPS_DOCUMENT_HANDLING_SELECT = "3088";

        private const int WIA_DATA_TYPE_COLOR = 3;
        private const int WIA_DATA_TYPE_GRAYSCALE = 2;
        private const int WIA_DATA_TYPE_BW = 0;

        private const int FEEDER = 1;
        private const int FLATBED = 2;

        public event EventHandler<ScanProgressEventArgs> ScanProgress;
        public event EventHandler<ScanCompleteEventArgs> ScanComplete;

        public List<ScannerInfo> GetAvailableScanners()
        {
            var scanners = new List<ScannerInfo>();

            try
            {
                // Create WIA DeviceManager using late binding
                Type deviceManagerType = Type.GetTypeFromProgID("WIA.DeviceManager");
                if (deviceManagerType == null)
                {
                    throw new ScanException("WIA not available on this system");
                }

                dynamic deviceManager = Activator.CreateInstance(deviceManagerType);

                foreach (dynamic deviceInfo in deviceManager.DeviceInfos)
                {
                    // WiaDeviceType.ScannerDeviceType = 1
                    if ((int)deviceInfo.Type == 1)
                    {
                        scanners.Add(new ScannerInfo
                        {
                            DeviceId = deviceInfo.DeviceID,
                            Name = GetDynamicPropertyValue(deviceInfo.Properties, "Name")?.ToString() ?? "Unknown Scanner",
                            Manufacturer = GetDynamicPropertyValue(deviceInfo.Properties, "Manufacturer")?.ToString() ?? ""
                        });
                    }
                }
            }
            catch (COMException ex)
            {
                throw new ScanException("Error getting scanner list: " + ex.Message, ex);
            }

            return scanners;
        }

        public ScannerInfo FindHPScanner()
        {
            var scanners = GetAvailableScanners();
            return scanners.Find(s =>
                s.Name.Contains("HP") || s.Name.Contains("LaserJet") ||
                s.Manufacturer.Contains("HP") || s.Manufacturer.Contains("Hewlett"));
        }

        public Image ScanPreview(ScanSettings settings)
        {
            var previewSettings = new ScanSettings
            {
                Source = settings.Source,
                Dpi = 75,
                ColorMode = settings.ColorMode
            };

            return ScanSinglePage(previewSettings);
        }

        public Image ScanSinglePage(ScanSettings settings)
        {
            try
            {
                var scanner = FindHPScanner();
                if (scanner == null)
                {
                    throw new ScanException("HP scanner not found");
                }

                Type deviceManagerType = Type.GetTypeFromProgID("WIA.DeviceManager");
                dynamic deviceManager = Activator.CreateInstance(deviceManagerType);
                dynamic device = null;

                foreach (dynamic deviceInfo in deviceManager.DeviceInfos)
                {
                    if (deviceInfo.DeviceID == scanner.DeviceId)
                    {
                        device = deviceInfo.Connect();
                        break;
                    }
                }

                if (device == null)
                {
                    throw new ScanException("Failed to connect to scanner");
                }

                // Select scan source
                dynamic item = device.Items[1];
                SelectScanSource(item, settings.Source);
                SetScannerSettings(item, settings);

                OnProgress(10, "Scanning...");

                // Perform scan - FormatID for BMP
                string wiaFormatBMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";
                dynamic imageFile = item.Transfer(wiaFormatBMP);

                OnProgress(80, "Processing image...");

                var image = ConvertToImage(imageFile);

                OnProgress(100, "Done");

                return image;
            }
            catch (COMException ex)
            {
                throw new ScanException("Scan error: " + GetWiaErrorMessage(ex.ErrorCode), ex);
            }
        }

        public List<Image> ScanMultiplePages(ScanSettings settings)
        {
            var pages = new List<Image>();

            if (settings.Source == ScanSource.ADF)
            {
                pages = ScanFromAdf(settings);
            }
            else
            {
                var page = ScanSinglePage(settings);
                if (page != null)
                {
                    pages.Add(page);
                }
            }

            return pages;
        }

        private List<Image> ScanFromAdf(ScanSettings settings)
        {
            var pages = new List<Image>();

            try
            {
                var scanner = FindHPScanner();
                if (scanner == null)
                {
                    throw new ScanException("HP scanner not found");
                }

                Type deviceManagerType = Type.GetTypeFromProgID("WIA.DeviceManager");
                dynamic deviceManager = Activator.CreateInstance(deviceManagerType);
                dynamic device = null;

                foreach (dynamic deviceInfo in deviceManager.DeviceInfos)
                {
                    if (deviceInfo.DeviceID == scanner.DeviceId)
                    {
                        device = deviceInfo.Connect();
                        break;
                    }
                }

                if (device == null)
                {
                    throw new ScanException("Failed to connect to scanner");
                }

                dynamic item = device.Items[1];
                SelectScanSource(item, ScanSource.ADF);
                SetScannerSettings(item, settings);

                int pageNumber = 0;
                bool hasMorePages = true;
                string wiaFormatBMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";

                while (hasMorePages)
                {
                    try
                    {
                        pageNumber++;
                        OnProgress((pageNumber * 10) % 90, $"Scanning page {pageNumber}...");

                        dynamic imageFile = item.Transfer(wiaFormatBMP);
                        var image = ConvertToImage(imageFile);
                        pages.Add(image);

                        hasMorePages = HasMorePagesInAdf(device);
                    }
                    catch (COMException ex)
                    {
                        // WIA_ERROR_PAPER_EMPTY
                        if (ex.ErrorCode == unchecked((int)0x80210003))
                        {
                            hasMorePages = false;
                        }
                        else
                        {
                            throw;
                        }
                    }
                }

                OnProgress(100, $"Scanned pages: {pages.Count}");
            }
            catch (COMException ex)
            {
                throw new ScanException("ADF scan error: " + GetWiaErrorMessage(ex.ErrorCode), ex);
            }

            return pages;
        }

        public void SaveScan(Image image, string filePath, Models.ImageFormat format)
        {
            switch (format)
            {
                case Models.ImageFormat.JPEG:
                    image.Save(filePath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    break;
                case Models.ImageFormat.PNG:
                    image.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                    break;
                case Models.ImageFormat.TIFF:
                    image.Save(filePath, System.Drawing.Imaging.ImageFormat.Tiff);
                    break;
                case Models.ImageFormat.PDF:
                    var pdfService = new PdfService();
                    pdfService.ImageToPdf(image, filePath);
                    break;
                default:
                    image.Save(filePath);
                    break;
            }
        }

        public void SaveMultipleScans(List<Image> images, string filePath, Models.ImageFormat format)
        {
            if (images == null || images.Count == 0) return;

            if (format == Models.ImageFormat.PDF)
            {
                var pdfService = new PdfService();
                pdfService.ImagesToPdf(images, filePath);
            }
            else if (format == Models.ImageFormat.TIFF)
            {
                SaveMultiPageTiff(images, filePath);
            }
            else
            {
                for (int i = 0; i < images.Count; i++)
                {
                    var pageFileName = Path.Combine(
                        Path.GetDirectoryName(filePath),
                        Path.GetFileNameWithoutExtension(filePath) + $"_{i + 1}" + Path.GetExtension(filePath));
                    SaveScan(images[i], pageFileName, format);
                }
            }
        }

        private void SaveMultiPageTiff(List<Image> images, string filePath)
        {
            if (images.Count == 0) return;

            var encoder = GetEncoder(System.Drawing.Imaging.ImageFormat.Tiff);
            var encoderParams = new EncoderParameters(1);

            using (var firstImage = (Image)images[0].Clone())
            {
                encoderParams.Param[0] = new EncoderParameter(
                    System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                firstImage.Save(filePath, encoder, encoderParams);

                encoderParams.Param[0] = new EncoderParameter(
                    System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);

                for (int i = 1; i < images.Count; i++)
                {
                    firstImage.SaveAdd(images[i], encoderParams);
                }

                encoderParams.Param[0] = new EncoderParameter(
                    System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);
                firstImage.SaveAdd(encoderParams);
            }
        }

        private ImageCodecInfo GetEncoder(System.Drawing.Imaging.ImageFormat format)
        {
            var codecs = ImageCodecInfo.GetImageEncoders();
            foreach (var codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                    return codec;
            }
            return null;
        }

        private void SelectScanSource(dynamic item, ScanSource source)
        {
            try
            {
                int sourceValue = source == ScanSource.ADF ? FEEDER : FLATBED;
                SetDynamicProperty(item.Properties, WIA_IPS_DOCUMENT_HANDLING_SELECT, sourceValue);
            }
            catch
            {
                // Use default if setting fails
            }
        }

        private void SetScannerSettings(dynamic item, ScanSettings settings)
        {
            try
            {
                // Set DPI
                SetDynamicProperty(item.Properties, WIA_IPS_XRES, settings.Dpi);
                SetDynamicProperty(item.Properties, WIA_IPS_YRES, settings.Dpi);

                // Set color mode
                int colorType;
                switch (settings.ColorMode)
                {
                    case Models.ColorMode.Color:
                        colorType = WIA_DATA_TYPE_COLOR;
                        break;
                    case Models.ColorMode.Grayscale:
                        colorType = WIA_DATA_TYPE_GRAYSCALE;
                        break;
                    case Models.ColorMode.BlackWhite:
                        colorType = WIA_DATA_TYPE_BW;
                        break;
                    default:
                        colorType = WIA_DATA_TYPE_COLOR;
                        break;
                }
                SetDynamicProperty(item.Properties, WIA_IPA_DATATYPE, colorType);

                // Set scan area (A4 at given DPI)
                int widthPixels = (int)(8.27 * settings.Dpi);
                int heightPixels = (int)(11.69 * settings.Dpi);
                SetDynamicProperty(item.Properties, WIA_IPS_XEXTENT, widthPixels);
                SetDynamicProperty(item.Properties, WIA_IPS_YEXTENT, heightPixels);
            }
            catch
            {
                // Ignore setting errors, use defaults
            }
        }

        private bool HasMorePagesInAdf(dynamic device)
        {
            try
            {
                var prop = GetDynamicPropertyValue(device.Properties, "Document Handling Status");
                if (prop != null)
                {
                    int status = Convert.ToInt32(prop);
                    return (status & 0x01) != 0; // FEED_READY
                }
            }
            catch
            {
            }
            return false;
        }

        private Image ConvertToImage(dynamic imageFile)
        {
            dynamic vector = imageFile.FileData;
            byte[] bytes = (byte[])vector.BinaryData;

            using (var ms = new MemoryStream(bytes))
            {
                return Image.FromStream(ms);
            }
        }

        private object GetDynamicPropertyValue(dynamic properties, string name)
        {
            foreach (dynamic prop in properties)
            {
                if (prop.Name == name)
                    return prop.Value;
            }
            return null;
        }

        private void SetDynamicProperty(dynamic properties, string propertyId, object value)
        {
            foreach (dynamic prop in properties)
            {
                if (prop.PropertyID.ToString() == propertyId)
                {
                    prop.Value = value;
                    return;
                }
            }
        }

        private string GetWiaErrorMessage(int errorCode)
        {
            switch (unchecked((uint)errorCode))
            {
                case 0x80210001: return "General WIA error";
                case 0x80210002: return "Paper jam";
                case 0x80210003: return "No paper in feeder";
                case 0x80210005: return "Device busy";
                case 0x80210006: return "Device offline";
                case 0x80210009: return "Invalid settings";
                case 0x8021000C: return "Scanner cover open";
                case 0x8021000D: return "Scanner lamp off";
                case 0x80210010: return "Device not found";
                default: return $"Error code: 0x{errorCode:X8}";
            }
        }

        protected virtual void OnProgress(int percent, string message)
        {
            ScanProgress?.Invoke(this, new ScanProgressEventArgs { Percent = percent, Message = message });
        }
    }

    public class ScannerInfo
    {
        public string DeviceId { get; set; }
        public string Name { get; set; }
        public string Manufacturer { get; set; }
    }

    public class ScanProgressEventArgs : EventArgs
    {
        public int Percent { get; set; }
        public string Message { get; set; }
    }

    public class ScanCompleteEventArgs : EventArgs
    {
        public Image ScannedImage { get; set; }
        public List<Image> ScannedImages { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class ScanException : Exception
    {
        public ScanException(string message) : base(message) { }
        public ScanException(string message, Exception inner) : base(message, inner) { }
    }
}
