using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace MFPControlCenter.Services
{
    /// <summary>
    /// Сервис для работы с документами (Word, PowerPoint, текстовые файлы)
    /// </summary>
    public class DocumentService
    {
        /// <summary>
        /// Получить список поддерживаемых расширений
        /// </summary>
        public static string[] SupportedExtensions => new[]
        {
            ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".gif",
            ".doc", ".docx", ".pptx", ".ppt", ".txt", ".rtf", ".xls", ".xlsx"
        };

        /// <summary>
        /// Проверить, поддерживается ли формат файла
        /// </summary>
        public static bool IsSupported(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLower();
            return Array.Exists(SupportedExtensions, e => e == ext);
        }

        /// <summary>
        /// Получить количество страниц в документе
        /// </summary>
        public int GetPageCount(string filePath)
        {
            var ext = Path.GetExtension(filePath).ToLower();

            switch (ext)
            {
                case ".txt":
                    return GetTextFilePageCount(filePath);
                case ".doc":
                case ".docx":
                    return GetWordPageCount(filePath);
                case ".ppt":
                case ".pptx":
                    return GetPowerPointPageCount(filePath);
                case ".pdf":
                    var pdfService = new PdfService();
                    return pdfService.GetPageCount(filePath);
                case ".jpg":
                case ".jpeg":
                case ".png":
                case ".bmp":
                case ".tiff":
                case ".tif":
                case ".gif":
                    return 1;
                default:
                    return 1;
            }
        }

        /// <summary>
        /// Получить превью страниц документа
        /// </summary>
        public List<Image> GetPagePreviews(string filePath, int dpi = 96)
        {
            var ext = Path.GetExtension(filePath).ToLower();
            var previews = new List<Image>();

            try
            {
                switch (ext)
                {
                    case ".txt":
                        previews = RenderTextFile(filePath, dpi);
                        break;
                    case ".doc":
                    case ".docx":
                        previews = RenderWordDocument(filePath, dpi);
                        break;
                    case ".ppt":
                    case ".pptx":
                        previews = RenderPowerPoint(filePath, dpi);
                        break;
                    case ".rtf":
                        previews = RenderRtfFile(filePath, dpi);
                        break;
                    case ".pdf":
                        previews = RenderPdfDocument(filePath, dpi);
                        break;
                    case ".jpg":
                    case ".jpeg":
                    case ".png":
                    case ".bmp":
                    case ".tiff":
                    case ".tif":
                    case ".gif":
                        previews.Add(Image.FromFile(filePath));
                        break;
                    default:
                        // Создаём заглушку с именем файла
                        previews.Add(CreatePlaceholderImage(filePath));
                        break;
                }
            }
            catch (Exception ex)
            {
                // При ошибке создаём заглушку с сообщением об ошибке
                previews.Add(CreateErrorImage(ex.Message));
            }

            return previews;
        }

        /// <summary>
        /// Печать документа через приложение по умолчанию
        /// </summary>
        public void PrintViaDefaultApp(string filePath, string printerName)
        {
            var ext = Path.GetExtension(filePath).ToLower();

            switch (ext)
            {
                case ".doc":
                case ".docx":
                    PrintWordDocument(filePath, printerName);
                    break;
                case ".ppt":
                case ".pptx":
                    PrintPowerPoint(filePath, printerName);
                    break;
                case ".xls":
                case ".xlsx":
                    PrintExcel(filePath, printerName);
                    break;
                default:
                    // Печать через shell
                    PrintViaShell(filePath, printerName);
                    break;
            }
        }

        #region Text Files

        private int GetTextFilePageCount(string filePath)
        {
            var text = File.ReadAllText(filePath, Encoding.UTF8);
            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            int linesPerPage = 50; // Примерно 50 строк на страницу
            return (int)Math.Ceiling((double)lines.Length / linesPerPage);
        }

        private List<Image> RenderTextFile(string filePath, int dpi)
        {
            var pages = new List<Image>();
            var text = File.ReadAllText(filePath, Encoding.UTF8);
            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            int linesPerPage = 50;
            var font = new Font("Consolas", 10);
            float lineHeight = font.GetHeight() + 2;

            // A4 размер в пикселях при заданном DPI
            int pageWidth = (int)(8.27 * dpi);
            int pageHeight = (int)(11.69 * dpi);
            int margin = (int)(0.5 * dpi); // 0.5 inch margin

            int currentLine = 0;
            while (currentLine < lines.Length)
            {
                var bitmap = new Bitmap(pageWidth, pageHeight);
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Color.White);
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                    float y = margin;
                    int pageLines = 0;

                    while (currentLine < lines.Length && pageLines < linesPerPage)
                    {
                        g.DrawString(lines[currentLine], font, Brushes.Black, margin, y);
                        y += lineHeight;
                        currentLine++;
                        pageLines++;
                    }
                }
                pages.Add(bitmap);
            }

            font.Dispose();

            if (pages.Count == 0)
            {
                pages.Add(CreatePlaceholderImage(filePath));
            }

            return pages;
        }

        #endregion

        #region PDF Documents

        private List<Image> RenderPdfDocument(string filePath, int dpi)
        {
            var pages = new List<Image>();

            try
            {
                var pdfService = new PdfService();
                pages = pdfService.PdfToImages(filePath, dpi);
            }
            catch (Exception ex)
            {
                pages.Add(CreateErrorImage($"Ошибка чтения PDF:\n{ex.Message}"));
            }

            if (pages.Count == 0)
            {
                pages.Add(CreatePlaceholderImage(filePath));
            }

            return pages;
        }

        #endregion

        #region Word Documents

        private int GetWordPageCount(string filePath)
        {
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null) return 1;

                dynamic wordApp = Activator.CreateInstance(wordType);
                wordApp.Visible = false;

                try
                {
                    dynamic doc = wordApp.Documents.Open(filePath, ReadOnly: true);
                    int pageCount = doc.ComputeStatistics(2); // wdStatisticPages = 2
                    doc.Close(false);
                    return pageCount;
                }
                finally
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
            catch
            {
                return 1;
            }
        }

        private List<Image> RenderWordDocument(string filePath, int dpi)
        {
            var pages = new List<Image>();

            try
            {
                // Попытка экспорта через Word в PDF, затем рендеринг PDF
                string tempPdf = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");

                if (ExportWordToPdf(filePath, tempPdf))
                {
                    try
                    {
                        // Используем PdfService для реального рендеринга PDF
                        var pdfService = new PdfService();
                        pages = pdfService.PdfToImages(tempPdf, dpi);
                    }
                    finally
                    {
                        try { File.Delete(tempPdf); } catch { }
                    }
                }
                else
                {
                    // Word не установлен - создаём заглушку
                    pages.Add(CreatePlaceholderImage(filePath, "Для предпросмотра Word документов\nтребуется Microsoft Word"));
                }
            }
            catch (Exception ex)
            {
                pages.Add(CreateErrorImage($"Ошибка чтения документа:\n{ex.Message}"));
            }

            if (pages.Count == 0)
            {
                pages.Add(CreatePlaceholderImage(filePath));
            }

            return pages;
        }

        private bool ExportWordToPdf(string wordPath, string pdfPath)
        {
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null) return false;

                dynamic wordApp = Activator.CreateInstance(wordType);
                wordApp.Visible = false;

                try
                {
                    dynamic doc = wordApp.Documents.Open(wordPath, ReadOnly: true);
                    doc.ExportAsFixedFormat(pdfPath, 17); // wdExportFormatPDF = 17
                    doc.Close(false);
                    return true;
                }
                finally
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
            catch
            {
                return false;
            }
        }

        private void PrintWordDocument(string filePath, string printerName)
        {
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    PrintViaShell(filePath, printerName);
                    return;
                }

                dynamic wordApp = Activator.CreateInstance(wordType);
                wordApp.Visible = false;

                try
                {
                    dynamic doc = wordApp.Documents.Open(filePath, ReadOnly: true);
                    wordApp.ActivePrinter = printerName;
                    doc.PrintOut();
                    doc.Close(false);
                }
                finally
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
            catch
            {
                PrintViaShell(filePath, printerName);
            }
        }

        #endregion

        #region PowerPoint

        private int GetPowerPointPageCount(string filePath)
        {
            try
            {
                Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null) return 1;

                dynamic pptApp = Activator.CreateInstance(pptType);

                try
                {
                    dynamic presentation = pptApp.Presentations.Open(filePath, ReadOnly: true, WithWindow: false);
                    int slideCount = presentation.Slides.Count;
                    presentation.Close();
                    return slideCount;
                }
                finally
                {
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
                }
            }
            catch
            {
                return 1;
            }
        }

        private List<Image> RenderPowerPoint(string filePath, int dpi)
        {
            var pages = new List<Image>();

            try
            {
                Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    pages.Add(CreatePlaceholderImage(filePath, "Для предпросмотра презентаций\nтребуется Microsoft PowerPoint"));
                    return pages;
                }

                dynamic pptApp = Activator.CreateInstance(pptType);

                try
                {
                    dynamic presentation = pptApp.Presentations.Open(filePath, ReadOnly: true, WithWindow: false);
                    int slideCount = presentation.Slides.Count;

                    // Экспорт слайдов как изображений
                    string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                    Directory.CreateDirectory(tempDir);

                    try
                    {
                        presentation.Export(tempDir, "PNG", 1024, 768);

                        // Загрузка экспортированных изображений
                        var slideFiles = Directory.GetFiles(tempDir, "*.PNG");
                        Array.Sort(slideFiles);

                        foreach (var slideFile in slideFiles)
                        {
                            using (var img = Image.FromFile(slideFile))
                            {
                                pages.Add(new Bitmap(img));
                            }
                        }
                    }
                    finally
                    {
                        try { Directory.Delete(tempDir, true); } catch { }
                    }

                    presentation.Close();
                }
                finally
                {
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
                }
            }
            catch (Exception ex)
            {
                pages.Add(CreateErrorImage($"Ошибка чтения презентации:\n{ex.Message}"));
            }

            if (pages.Count == 0)
            {
                pages.Add(CreatePlaceholderImage(filePath));
            }

            return pages;
        }

        private void PrintPowerPoint(string filePath, string printerName)
        {
            try
            {
                Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    PrintViaShell(filePath, printerName);
                    return;
                }

                dynamic pptApp = Activator.CreateInstance(pptType);

                try
                {
                    dynamic presentation = pptApp.Presentations.Open(filePath, ReadOnly: true, WithWindow: false);
                    presentation.PrintOptions.ActivePrinter = printerName;
                    presentation.PrintOut();
                    presentation.Close();
                }
                finally
                {
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
                }
            }
            catch
            {
                PrintViaShell(filePath, printerName);
            }
        }

        #endregion

        #region Excel

        private void PrintExcel(string filePath, string printerName)
        {
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    PrintViaShell(filePath, printerName);
                    return;
                }

                dynamic excelApp = Activator.CreateInstance(excelType);
                excelApp.Visible = false;

                try
                {
                    dynamic workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);
                    excelApp.ActivePrinter = printerName;
                    workbook.PrintOut();
                    workbook.Close(false);
                }
                finally
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch
            {
                PrintViaShell(filePath, printerName);
            }
        }

        #endregion

        #region RTF

        private List<Image> RenderRtfFile(string filePath, int dpi)
        {
            // RTF можно обработать через RichTextBox
            var pages = new List<Image>();

            try
            {
                var rtfText = File.ReadAllText(filePath);
                // Упрощённый рендеринг - как текстовый файл
                pages = RenderTextFile(filePath, dpi);
            }
            catch
            {
                pages.Add(CreatePlaceholderImage(filePath));
            }

            return pages;
        }

        #endregion

        #region Shell Printing

        private void PrintViaShell(string filePath, string printerName)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = filePath,
                    Verb = "print",
                    UseShellExecute = true,
                    CreateNoWindow = true
                };

                // Установка принтера по умолчанию временно
                SetDefaultPrinter(printerName);

                using (var process = Process.Start(startInfo))
                {
                    process?.WaitForExit(30000); // Ждём максимум 30 секунд
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Не удалось напечатать файл: {ex.Message}", ex);
            }
        }

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetDefaultPrinter(string printerName);

        #endregion

        #region Placeholder Images

        private Image CreatePlaceholderImage(string filePath, string additionalText = null)
        {
            int width = 794;  // A4 width at 96 DPI
            int height = 1123; // A4 height at 96 DPI

            var bitmap = new Bitmap(width, height);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                var fileName = Path.GetFileName(filePath);
                var titleFont = new Font("Segoe UI", 16, FontStyle.Bold);
                var textFont = new Font("Segoe UI", 12);

                // Иконка документа (простой прямоугольник)
                int iconX = width / 2 - 40;
                int iconY = height / 2 - 100;
                g.FillRectangle(Brushes.LightGray, iconX, iconY, 80, 100);
                g.DrawRectangle(Pens.Gray, iconX, iconY, 80, 100);

                // Имя файла
                var titleSize = g.MeasureString(fileName, titleFont);
                g.DrawString(fileName, titleFont, Brushes.Black,
                    (width - titleSize.Width) / 2, iconY + 120);

                // Дополнительный текст
                if (!string.IsNullOrEmpty(additionalText))
                {
                    var textSize = g.MeasureString(additionalText, textFont);
                    g.DrawString(additionalText, textFont, Brushes.Gray,
                        (width - textSize.Width) / 2, iconY + 160);
                }

                titleFont.Dispose();
                textFont.Dispose();
            }

            return bitmap;
        }

        private Image CreatePagePlaceholder(string pageText, int totalPages, string filePath, int dpi)
        {
            int width = (int)(8.27 * dpi);
            int height = (int)(11.69 * dpi);

            var bitmap = new Bitmap(width, height);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                var font = new Font("Segoe UI", 14);
                var smallFont = new Font("Segoe UI", 10);

                // Рамка страницы
                g.DrawRectangle(new Pen(Color.LightGray, 2), 20, 20, width - 40, height - 40);

                // Текст страницы
                var textSize = g.MeasureString(pageText, font);
                g.DrawString(pageText, font, Brushes.Black,
                    (width - textSize.Width) / 2, height / 2 - 20);

                // Имя файла внизу
                var fileName = Path.GetFileName(filePath);
                var fileSize = g.MeasureString(fileName, smallFont);
                g.DrawString(fileName, smallFont, Brushes.Gray,
                    (width - fileSize.Width) / 2, height - 60);

                font.Dispose();
                smallFont.Dispose();
            }

            return bitmap;
        }

        private Image CreateErrorImage(string errorMessage)
        {
            int width = 794;
            int height = 1123;

            var bitmap = new Bitmap(width, height);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                var font = new Font("Segoe UI", 12);

                // Красная иконка ошибки
                g.FillEllipse(Brushes.Red, width / 2 - 30, height / 2 - 100, 60, 60);
                g.DrawString("!", new Font("Segoe UI", 30, FontStyle.Bold), Brushes.White,
                    width / 2 - 12, height / 2 - 95);

                // Текст ошибки
                var textSize = g.MeasureString(errorMessage, font);
                g.DrawString(errorMessage, font, Brushes.Red,
                    (width - textSize.Width) / 2, height / 2);

                font.Dispose();
            }

            return bitmap;
        }

        #endregion
    }
}
