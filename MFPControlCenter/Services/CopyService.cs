using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using MFPControlCenter.Models;

namespace MFPControlCenter.Services
{
    public class CopyService
    {
        private readonly ScanService _scanService;
        private readonly PrintService _printService;
        private readonly PdfService _pdfService;

        public event EventHandler<CopyProgressEventArgs> CopyProgress;

        public CopyService()
        {
            _scanService = new ScanService();
            _printService = new PrintService();
            _pdfService = new PdfService();
        }

        public void InstantCopy(CopySettings settings)
        {
            OnProgress(0, "Начало копирования...");

            var scanSettings = CreateScanSettings(settings);
            var printSettings = CreatePrintSettings(settings);

            if (settings.Source == ScanSource.ADF)
            {
                // Сканирование и печать из ADF постранично
                InstantCopyFromAdf(scanSettings, printSettings, settings.Copies);
            }
            else
            {
                // Одиночное сканирование и печать
                OnProgress(10, "Сканирование...");
                var image = _scanService.ScanSinglePage(scanSettings);

                if (image != null)
                {
                    // Применение настроек яркости/контрастности
                    if (settings.Brightness != 0 || settings.Contrast != 0)
                    {
                        image = AdjustImage(image, settings.Brightness, settings.Contrast);
                    }

                    // Масштабирование
                    if (settings.ScalePercent != 100)
                    {
                        image = ScaleImage(image, settings.ScalePercent);
                    }

                    OnProgress(50, "Печать...");
                    _printService.PrintImage(image, printSettings);
                    image.Dispose();
                }
            }

            OnProgress(100, "Копирование завершено");
        }

        private void InstantCopyFromAdf(ScanSettings scanSettings, PrintSettings printSettings, int copies)
        {
            bool hasMorePages = true;
            int pageNumber = 0;

            while (hasMorePages)
            {
                try
                {
                    pageNumber++;
                    OnProgress((pageNumber * 20) % 80, $"Страница {pageNumber}: сканирование...");

                    var image = _scanService.ScanSinglePage(scanSettings);

                    if (image != null)
                    {
                        OnProgress((pageNumber * 20) % 80 + 10, $"Страница {pageNumber}: печать...");
                        printSettings.Copies = copies;
                        _printService.PrintImage(image, printSettings);
                        image.Dispose();
                    }
                    else
                    {
                        hasMorePages = false;
                    }
                }
                catch (ScanException ex)
                {
                    // Если закончились страницы в ADF
                    if (ex.Message.Contains("Нет бумаги"))
                    {
                        hasMorePages = false;
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        public void DeferredCopy(CopySettings settings)
        {
            OnProgress(0, "Начало отложенного копирования...");

            var scanSettings = CreateScanSettings(settings);
            var printSettings = CreatePrintSettings(settings);

            OnProgress(10, "Сканирование всех страниц...");

            // Сканирование всех страниц
            List<Image> pages;
            if (settings.Source == ScanSource.ADF)
            {
                pages = _scanService.ScanMultiplePages(scanSettings);
            }
            else
            {
                // Для планшета - сканируем одну страницу
                var singlePage = _scanService.ScanSinglePage(scanSettings);
                pages = new List<Image> { singlePage };
            }

            if (pages.Count == 0)
            {
                OnProgress(100, "Нет страниц для копирования");
                return;
            }

            OnProgress(50, $"Отсканировано {pages.Count} страниц. Обработка...");

            // Обработка изображений
            for (int i = 0; i < pages.Count; i++)
            {
                if (settings.Brightness != 0 || settings.Contrast != 0)
                {
                    pages[i] = AdjustImage(pages[i], settings.Brightness, settings.Contrast);
                }

                if (settings.ScalePercent != 100)
                {
                    pages[i] = ScaleImage(pages[i], settings.ScalePercent);
                }
            }

            OnProgress(70, "Создание PDF...");

            // Создание временного PDF
            var tempPdfPath = Path.Combine(Path.GetTempPath(), $"copy_{Guid.NewGuid()}.pdf");
            _pdfService.ImagesToPdf(pages, tempPdfPath);

            OnProgress(85, "Печать документа...");

            // Печать PDF
            printSettings.Copies = settings.Copies;
            _printService.PrintFile(tempPdfPath, printSettings);

            // Очистка
            foreach (var page in pages)
            {
                page.Dispose();
            }

            try
            {
                File.Delete(tempPdfPath);
            }
            catch { }

            OnProgress(100, "Копирование завершено");
        }

        public void IdCopy(CopySettings settings, Action<string> promptCallback)
        {
            OnProgress(0, "ID-копирование: сканирование лицевой стороны...");

            var scanSettings = new ScanSettings
            {
                Source = ScanSource.Flatbed, // ID-копия всегда с планшета
                Dpi = 300,
                ColorMode = ColorMode.Color
            };

            // Сканирование лицевой стороны
            OnProgress(10, "Сканирование лицевой стороны...");
            var frontSide = _scanService.ScanSinglePage(scanSettings);

            if (frontSide == null)
            {
                OnProgress(100, "Ошибка: не удалось отсканировать лицевую сторону");
                return;
            }

            OnProgress(40, "Лицевая сторона отсканирована");

            // Запрос на переворот документа
            promptCallback?.Invoke("Переверните документ и нажмите OK для сканирования оборотной стороны");

            // Сканирование оборотной стороны
            OnProgress(50, "Сканирование оборотной стороны...");
            var backSide = _scanService.ScanSinglePage(scanSettings);

            if (backSide == null)
            {
                frontSide.Dispose();
                OnProgress(100, "Ошибка: не удалось отсканировать оборотную сторону");
                return;
            }

            OnProgress(70, "Создание ID-копии...");

            // Создание PDF с обеими сторонами на одном листе
            var tempPdfPath = Path.Combine(Path.GetTempPath(), $"idcopy_{Guid.NewGuid()}.pdf");
            _pdfService.CreateIdCopyPdf(frontSide, backSide, tempPdfPath);

            OnProgress(85, "Печать ID-копии...");

            // Печать
            var printSettings = CreatePrintSettings(settings);
            printSettings.Copies = settings.Copies;
            _printService.PrintFile(tempPdfPath, printSettings);

            // Очистка
            frontSide.Dispose();
            backSide.Dispose();

            try
            {
                File.Delete(tempPdfPath);
            }
            catch { }

            OnProgress(100, "ID-копирование завершено");
        }

        private ScanSettings CreateScanSettings(CopySettings copySettings)
        {
            return new ScanSettings
            {
                Source = copySettings.Source,
                Dpi = 300,
                ColorMode = ColorMode.Grayscale // Для копирования обычно достаточно градаций серого
            };
        }

        private PrintSettings CreatePrintSettings(CopySettings copySettings)
        {
            return new PrintSettings
            {
                PrinterName = _printService.FindHPLaserJetPrinter(),
                IsDuplex = copySettings.IsDuplex,
                Copies = copySettings.Copies
            };
        }

        private Image AdjustImage(Image original, int brightness, int contrast)
        {
            var bitmap = new Bitmap(original.Width, original.Height);

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

            original.Dispose();
            return bitmap;
        }

        private Image ScaleImage(Image original, int scalePercent)
        {
            int newWidth = (int)(original.Width * scalePercent / 100.0);
            int newHeight = (int)(original.Height * scalePercent / 100.0);

            var bitmap = new Bitmap(newWidth, newHeight);

            using (var g = Graphics.FromImage(bitmap))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(original, 0, 0, newWidth, newHeight);
            }

            original.Dispose();
            return bitmap;
        }

        protected virtual void OnProgress(int percent, string message)
        {
            CopyProgress?.Invoke(this, new CopyProgressEventArgs { Percent = percent, Message = message });
        }
    }

    public class CopyProgressEventArgs : EventArgs
    {
        public int Percent { get; set; }
        public string Message { get; set; }
    }
}
