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
        // WIA константы
        private const string WIA_DEVICE_PROPERTY_MANUFACTURER = "Manufacturer";
        private const string WIA_IPA_DATATYPE = "6146"; // Color/Grayscale/BW
        private const string WIA_IPA_DEPTH = "4104";    // Bit depth
        private const string WIA_IPS_XRES = "6147";     // Horizontal resolution
        private const string WIA_IPS_YRES = "6148";     // Vertical resolution
        private const string WIA_IPS_XEXTENT = "6151";  // Width in pixels
        private const string WIA_IPS_YEXTENT = "6152";  // Height in pixels
        private const string WIA_IPS_DOCUMENT_HANDLING_SELECT = "3088"; // Feeder/Flatbed

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
                var deviceManager = new WIA.DeviceManager();

                for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
                {
                    var deviceInfo = deviceManager.DeviceInfos[i];

                    if (deviceInfo.Type == WIA.WiaDeviceType.ScannerDeviceType)
                    {
                        scanners.Add(new ScannerInfo
                        {
                            DeviceId = deviceInfo.DeviceID,
                            Name = GetPropertyValue(deviceInfo.Properties, "Name")?.ToString() ?? "Unknown Scanner",
                            Manufacturer = GetPropertyValue(deviceInfo.Properties, "Manufacturer")?.ToString() ?? ""
                        });
                    }
                }
            }
            catch (COMException ex)
            {
                throw new ScanException("Ошибка при получении списка сканеров: " + ex.Message, ex);
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
            // Быстрое сканирование с низким разрешением для предпросмотра
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
                    throw new ScanException("HP сканер не найден");
                }

                var deviceManager = new WIA.DeviceManager();
                WIA.Device device = null;

                for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
                {
                    if (deviceManager.DeviceInfos[i].DeviceID == scanner.DeviceId)
                    {
                        device = deviceManager.DeviceInfos[i].Connect();
                        break;
                    }
                }

                if (device == null)
                {
                    throw new ScanException("Не удалось подключиться к сканеру");
                }

                // Выбор источника (планшет или ADF)
                var item = SelectScanSource(device, settings.Source);

                // Настройка параметров сканирования
                SetScannerSettings(item, settings);

                OnProgress(10, "Сканирование...");

                // Выполнение сканирования
                var imageFile = (WIA.ImageFile)item.Transfer(WIA.FormatID.wiaFormatBMP);

                OnProgress(80, "Обработка изображения...");

                // Конвертация в Image
                var image = ConvertToImage(imageFile);

                OnProgress(100, "Готово");

                return image;
            }
            catch (COMException ex)
            {
                throw new ScanException("Ошибка сканирования: " + GetWiaErrorMessage(ex.ErrorCode), ex);
            }
        }

        public List<Image> ScanMultiplePages(ScanSettings settings)
        {
            var pages = new List<Image>();

            if (settings.Source == ScanSource.ADF)
            {
                // Сканирование из автоподатчика
                pages = ScanFromAdf(settings);
            }
            else
            {
                // Для планшета сканируем одну страницу
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
                    throw new ScanException("HP сканер не найден");
                }

                var deviceManager = new WIA.DeviceManager();
                WIA.Device device = null;

                for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
                {
                    if (deviceManager.DeviceInfos[i].DeviceID == scanner.DeviceId)
                    {
                        device = deviceManager.DeviceInfos[i].Connect();
                        break;
                    }
                }

                if (device == null)
                {
                    throw new ScanException("Не удалось подключиться к сканеру");
                }

                var item = SelectScanSource(device, ScanSource.ADF);
                SetScannerSettings(item, settings);

                int pageNumber = 0;
                bool hasMorePages = true;

                while (hasMorePages)
                {
                    try
                    {
                        pageNumber++;
                        OnProgress((pageNumber * 10) % 90, $"Сканирование страницы {pageNumber}...");

                        var imageFile = (WIA.ImageFile)item.Transfer(WIA.FormatID.wiaFormatBMP);
                        var image = ConvertToImage(imageFile);
                        pages.Add(image);

                        // Проверка наличия следующей страницы
                        hasMorePages = HasMorePagesInAdf(device);
                    }
                    catch (COMException ex)
                    {
                        // WIA_ERROR_PAPER_EMPTY означает, что в ADF больше нет страниц
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

                OnProgress(100, $"Отсканировано страниц: {pages.Count}");
            }
            catch (COMException ex)
            {
                throw new ScanException("Ошибка при сканировании из ADF: " + GetWiaErrorMessage(ex.ErrorCode), ex);
            }

            return pages;
        }

        public void SaveScan(Image image, string filePath, ImageFormat format)
        {
            var extension = Path.GetExtension(filePath).ToLower();

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

        public void SaveMultipleScans(List<Image> images, string filePath, ImageFormat format)
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
                // Для других форматов сохраняем каждую страницу отдельно
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

        private WIA.Item SelectScanSource(WIA.Device device, ScanSource source)
        {
            WIA.Item item = device.Items[1];

            try
            {
                // Попытка выбрать источник
                int sourceValue = source == ScanSource.ADF ? FEEDER : FLATBED;
                SetProperty(item.Properties, WIA_IPS_DOCUMENT_HANDLING_SELECT, sourceValue);
            }
            catch
            {
                // Если не удалось установить источник, используем по умолчанию
            }

            return item;
        }

        private void SetScannerSettings(WIA.Item item, ScanSettings settings)
        {
            try
            {
                // Установка DPI
                SetProperty(item.Properties, WIA_IPS_XRES, settings.Dpi);
                SetProperty(item.Properties, WIA_IPS_YRES, settings.Dpi);

                // Установка цветового режима
                int colorType;
                switch (settings.ColorMode)
                {
                    case ColorMode.Color:
                        colorType = WIA_DATA_TYPE_COLOR;
                        break;
                    case ColorMode.Grayscale:
                        colorType = WIA_DATA_TYPE_GRAYSCALE;
                        break;
                    case ColorMode.BlackWhite:
                        colorType = WIA_DATA_TYPE_BW;
                        break;
                    default:
                        colorType = WIA_DATA_TYPE_COLOR;
                        break;
                }
                SetProperty(item.Properties, WIA_IPA_DATATYPE, colorType);

                // Установка области сканирования (A4 при заданном DPI)
                int widthPixels = (int)(8.27 * settings.Dpi);  // A4 width in inches
                int heightPixels = (int)(11.69 * settings.Dpi); // A4 height in inches
                SetProperty(item.Properties, WIA_IPS_XEXTENT, widthPixels);
                SetProperty(item.Properties, WIA_IPS_YEXTENT, heightPixels);
            }
            catch
            {
                // Игнорируем ошибки настройки, используем значения по умолчанию
            }
        }

        private bool HasMorePagesInAdf(WIA.Device device)
        {
            try
            {
                var prop = GetPropertyValue(device.Properties, "Document Handling Status");
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

        private Image ConvertToImage(WIA.ImageFile imageFile)
        {
            var vector = imageFile.FileData;
            var bytes = (byte[])vector.get_BinaryData();

            using (var ms = new MemoryStream(bytes))
            {
                return Image.FromStream(ms);
            }
        }

        private object GetPropertyValue(WIA.IProperties properties, string name)
        {
            foreach (WIA.Property prop in properties)
            {
                if (prop.Name == name)
                    return prop.get_Value();
            }
            return null;
        }

        private void SetProperty(WIA.IProperties properties, string propertyId, object value)
        {
            foreach (WIA.Property prop in properties)
            {
                if (prop.PropertyID.ToString() == propertyId)
                {
                    prop.set_Value(value);
                    return;
                }
            }
        }

        private string GetWiaErrorMessage(int errorCode)
        {
            switch (unchecked((uint)errorCode))
            {
                case 0x80210001: return "Общая ошибка WIA";
                case 0x80210002: return "Бумага замята";
                case 0x80210003: return "Нет бумаги в податчике";
                case 0x80210005: return "Устройство занято";
                case 0x80210006: return "Устройство отключено";
                case 0x80210009: return "Некорректные настройки";
                case 0x8021000C: return "Крышка сканера открыта";
                case 0x8021000D: return "Лампа сканера выключена";
                case 0x80210010: return "Устройство не найдено";
                default: return $"Код ошибки: 0x{errorCode:X8}";
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
