using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Printing;
using MFPControlCenter.Models;

namespace MFPControlCenter.Services
{
    public class PrintService
    {
        private const string HP_PRINTER_NAME_PATTERN = "HP LaserJet";
        private PrintDocument _printDocument;
        private List<Image> _pagesToPrint;
        private int _currentPageIndex;

        public string FindHPLaserJetPrinter()
        {
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                if (printer.Contains(HP_PRINTER_NAME_PATTERN) || printer.Contains("1536"))
                {
                    return printer;
                }
            }
            return null;
        }

        public List<string> GetAvailablePrinters()
        {
            var printers = new List<string>();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                printers.Add(printer);
            }
            return printers;
        }

        public bool IsPrinterOnline(string printerName)
        {
            try
            {
                using (var server = new LocalPrintServer())
                {
                    var queue = server.GetPrintQueues()
                        .FirstOrDefault(q => q.Name.Equals(printerName, StringComparison.OrdinalIgnoreCase));

                    if (queue != null)
                    {
                        return !queue.IsOffline;
                    }
                }
            }
            catch
            {
                // Игнорируем ошибки
            }
            return false;
        }

        public PrinterCapabilities GetPrinterCapabilities(string printerName)
        {
            var settings = new PrinterSettings { PrinterName = printerName };
            var capabilities = new PrinterCapabilities
            {
                SupportsDuplex = settings.CanDuplex,
                SupportsColor = settings.SupportsColor,
                MaxCopies = settings.MaximumCopies
            };

            foreach (System.Drawing.Printing.PaperSize size in settings.PaperSizes)
            {
                capabilities.PaperSizes.Add(size.PaperName);
            }

            return capabilities;
        }

        public void PrintImage(Image image, PrintSettings settings)
        {
            PrintImages(new List<Image> { image }, settings);
        }

        public void PrintImages(List<Image> images, PrintSettings settings)
        {
            // Apply image adjustments if any
            if (settings.Brightness != 0 || settings.Contrast != 0 || settings.Sharpness > 0)
            {
                _pagesToPrint = new List<Image>();
                foreach (var img in images)
                {
                    var adjusted = ApplyImageAdjustments(img, settings.Brightness, settings.Contrast, settings.Sharpness);
                    _pagesToPrint.Add(adjusted);
                }
            }
            else
            {
                _pagesToPrint = images;
            }

            _currentPageIndex = 0;

            _printDocument = new PrintDocument();
            _printDocument.PrinterSettings.PrinterName = settings.PrinterName ?? FindHPLaserJetPrinter();
            _printDocument.PrinterSettings.Copies = (short)settings.Copies;

            if (settings.IsDuplex && _printDocument.PrinterSettings.CanDuplex)
            {
                _printDocument.PrinterSettings.Duplex = Duplex.Vertical;
            }

            ApplyPaperSize(_printDocument, settings.PaperSize);
            ApplyOrientation(_printDocument, settings.Orientation);

            _printDocument.PrintPage += PrintDocument_PrintPage;
            _printDocument.Print();
        }

        public void PrintFile(string filePath, PrintSettings settings)
        {
            var extension = Path.GetExtension(filePath).ToLower();

            // Изображения
            if (new[] { ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".gif" }.Contains(extension))
            {
                using (var image = Image.FromFile(filePath))
                {
                    PrintImage(image, settings);
                }
            }
            // PDF
            else if (extension == ".pdf")
            {
                PrintPdf(filePath, settings);
            }
            // Текстовые файлы
            else if (extension == ".txt")
            {
                PrintTextFile(filePath, settings);
            }
            // Документы Office (Word, PowerPoint, Excel)
            else if (new[] { ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".rtf" }.Contains(extension))
            {
                PrintOfficeDocument(filePath, settings);
            }
            else
            {
                throw new NotSupportedException($"Формат файла {extension} не поддерживается");
            }
        }

        private void PrintTextFile(string filePath, PrintSettings settings)
        {
            var documentService = new DocumentService();
            var pages = documentService.GetPagePreviews(filePath);

            if (pages.Count > 0)
            {
                var pagesToPrint = FilterPages(pages, settings.PageRange);
                PrintImages(pagesToPrint, settings);

                foreach (var page in pages)
                {
                    page.Dispose();
                }
            }
        }

        private void PrintOfficeDocument(string filePath, PrintSettings settings)
        {
            var documentService = new DocumentService();
            var printerName = settings.PrinterName ?? FindHPLaserJetPrinter();

            // Печать через COM-объекты Office
            documentService.PrintViaDefaultApp(filePath, printerName);
        }

        private void PrintPdf(string pdfPath, PrintSettings settings)
        {
            // Для печати PDF используем PdfService для конвертации в изображения
            var pdfService = new PdfService();
            var images = pdfService.PdfToImages(pdfPath);

            // Фильтрация страниц по диапазону
            var pagesToPrint = FilterPages(images, settings.PageRange);

            PrintImages(pagesToPrint, settings);

            // Освобождаем ресурсы
            foreach (var img in images)
            {
                img.Dispose();
            }
        }

        private List<Image> FilterPages(List<Image> allPages, string pageRange)
        {
            if (string.IsNullOrEmpty(pageRange) || pageRange.ToLower() == "all")
            {
                return allPages;
            }

            var selectedIndexes = ParsePageRange(pageRange, allPages.Count);
            return selectedIndexes.Select(i => allPages[i]).ToList();
        }

        private List<int> ParsePageRange(string range, int totalPages)
        {
            var result = new List<int>();
            var parts = range.Split(',');

            foreach (var part in parts)
            {
                if (part.Contains("-"))
                {
                    var rangeParts = part.Split('-');
                    if (rangeParts.Length == 2 &&
                        int.TryParse(rangeParts[0].Trim(), out int start) &&
                        int.TryParse(rangeParts[1].Trim(), out int end))
                    {
                        for (int i = start - 1; i < end && i < totalPages; i++)
                        {
                            if (i >= 0 && !result.Contains(i))
                                result.Add(i);
                        }
                    }
                }
                else if (int.TryParse(part.Trim(), out int page))
                {
                    int index = page - 1;
                    if (index >= 0 && index < totalPages && !result.Contains(index))
                        result.Add(index);
                }
            }

            result.Sort();
            return result;
        }

        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (_currentPageIndex < _pagesToPrint.Count)
            {
                var image = _pagesToPrint[_currentPageIndex];

                // Масштабирование изображения под страницу
                var destRect = CalculateFitRectangle(
                    e.MarginBounds,
                    new Size(image.Width, image.Height));

                e.Graphics.DrawImage(image, destRect);

                _currentPageIndex++;
                e.HasMorePages = _currentPageIndex < _pagesToPrint.Count;
            }
        }

        private Rectangle CalculateFitRectangle(Rectangle bounds, Size imageSize)
        {
            float ratioX = (float)bounds.Width / imageSize.Width;
            float ratioY = (float)bounds.Height / imageSize.Height;
            float ratio = Math.Min(ratioX, ratioY);

            int newWidth = (int)(imageSize.Width * ratio);
            int newHeight = (int)(imageSize.Height * ratio);

            int x = bounds.X + (bounds.Width - newWidth) / 2;
            int y = bounds.Y + (bounds.Height - newHeight) / 2;

            return new Rectangle(x, y, newWidth, newHeight);
        }

        private void ApplyPaperSize(PrintDocument doc, Models.PaperSize size)
        {
            string paperName;
            switch (size)
            {
                case Models.PaperSize.A4: paperName = "A4"; break;
                case Models.PaperSize.Letter: paperName = "Letter"; break;
                case Models.PaperSize.Legal: paperName = "Legal"; break;
                case Models.PaperSize.A5: paperName = "A5"; break;
                case Models.PaperSize.B5: paperName = "B5"; break;
                default: paperName = "A4"; break;
            }

            foreach (System.Drawing.Printing.PaperSize ps in doc.PrinterSettings.PaperSizes)
            {
                if (ps.PaperName.Contains(paperName))
                {
                    doc.DefaultPageSettings.PaperSize = ps;
                    break;
                }
            }
        }

        private void ApplyOrientation(PrintDocument doc, Models.Orientation orientation)
        {
            doc.DefaultPageSettings.Landscape = orientation == Models.Orientation.Landscape;
        }

        private Image ApplyImageAdjustments(Image original, int brightness, int contrast, int sharpness)
        {
            var bitmap = new Bitmap(original.Width, original.Height);

            // Apply brightness and contrast
            float brightnessFactor = 1.0f + brightness / 100.0f;
            float contrastFactor = 1.0f + contrast / 100.0f;

            float[][] matrix = {
                new float[] { contrastFactor, 0, 0, 0, 0 },
                new float[] { 0, contrastFactor, 0, 0, 0 },
                new float[] { 0, 0, contrastFactor, 0, 0 },
                new float[] { 0, 0, 0, 1, 0 },
                new float[] { brightnessFactor - 1, brightnessFactor - 1, brightnessFactor - 1, 0, 1 }
            };

            var colorMatrix = new System.Drawing.Imaging.ColorMatrix(matrix);
            var attributes = new System.Drawing.Imaging.ImageAttributes();
            attributes.SetColorMatrix(colorMatrix);

            using (var g = Graphics.FromImage(bitmap))
            {
                g.DrawImage(original,
                    new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                    0, 0, original.Width, original.Height,
                    GraphicsUnit.Pixel, attributes);
            }

            // Apply sharpness if needed
            if (sharpness > 0)
            {
                bitmap = ApplySharpness(bitmap, sharpness);
            }

            return bitmap;
        }

        private Bitmap ApplySharpness(Bitmap image, int amount)
        {
            // Simple sharpening using unsharp mask kernel
            float factor = amount / 100.0f;
            float[,] kernel = {
                { 0, -factor, 0 },
                { -factor, 1 + 4 * factor, -factor },
                { 0, -factor, 0 }
            };

            var result = new Bitmap(image.Width, image.Height);

            for (int x = 1; x < image.Width - 1; x++)
            {
                for (int y = 1; y < image.Height - 1; y++)
                {
                    float r = 0, g = 0, b = 0;

                    for (int kx = -1; kx <= 1; kx++)
                    {
                        for (int ky = -1; ky <= 1; ky++)
                        {
                            var pixel = image.GetPixel(x + kx, y + ky);
                            float kVal = kernel[kx + 1, ky + 1];
                            r += pixel.R * kVal;
                            g += pixel.G * kVal;
                            b += pixel.B * kVal;
                        }
                    }

                    r = Math.Max(0, Math.Min(255, r));
                    g = Math.Max(0, Math.Min(255, g));
                    b = Math.Max(0, Math.Min(255, b));

                    result.SetPixel(x, y, Color.FromArgb((int)r, (int)g, (int)b));
                }
            }

            image.Dispose();
            return result;
        }
    }

    public class PrinterCapabilities
    {
        public bool SupportsDuplex { get; set; }
        public bool SupportsColor { get; set; }
        public int MaxCopies { get; set; }
        public List<string> PaperSizes { get; set; } = new List<string>();
    }
}
