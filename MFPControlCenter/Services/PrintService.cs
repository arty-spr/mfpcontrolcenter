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
            _pagesToPrint = images;
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
    }

    public class PrinterCapabilities
    {
        public bool SupportsDuplex { get; set; }
        public bool SupportsColor { get; set; }
        public int MaxCopies { get; set; }
        public List<string> PaperSizes { get; set; } = new List<string>();
    }
}
