using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace MFPControlCenter.Services
{
    public class PdfService
    {
        public void ImageToPdf(Image image, string outputPath)
        {
            ImagesToPdf(new List<Image> { image }, outputPath);
        }

        public void ImagesToPdf(List<Image> images, string outputPath)
        {
            using (var document = new PdfDocument())
            {
                document.Info.Title = "Scanned Document";
                document.Info.Creator = "MFP Control Center";

                foreach (var image in images)
                {
                    var page = document.AddPage();

                    // Установка размера страницы A4
                    page.Width = XUnit.FromMillimeter(210);
                    page.Height = XUnit.FromMillimeter(297);

                    using (var gfx = XGraphics.FromPdfPage(page))
                    {
                        // Конвертация Image в XImage
                        using (var ms = new MemoryStream())
                        {
                            image.Save(ms, ImageFormat.Png);
                            ms.Position = 0;

                            using (var xImage = XImage.FromStream(ms))
                            {
                                // Масштабирование изображения под страницу с сохранением пропорций
                                double ratioX = page.Width.Point / xImage.PixelWidth;
                                double ratioY = page.Height.Point / xImage.PixelHeight;
                                double ratio = Math.Min(ratioX, ratioY);

                                double newWidth = xImage.PixelWidth * ratio;
                                double newHeight = xImage.PixelHeight * ratio;

                                double x = (page.Width.Point - newWidth) / 2;
                                double y = (page.Height.Point - newHeight) / 2;

                                gfx.DrawImage(xImage, x, y, newWidth, newHeight);
                            }
                        }
                    }
                }

                document.Save(outputPath);
            }
        }

        public List<Image> PdfToImages(string pdfPath)
        {
            var images = new List<Image>();

            // PDFsharp не поддерживает рендеринг PDF в изображения напрямую
            // Для этого нужна дополнительная библиотека типа PdfiumViewer
            // Здесь реализуем базовую заглушку

            throw new NotImplementedException(
                "Для конвертации PDF в изображения требуется дополнительная библиотека. " +
                "Рекомендуется установить PdfiumViewer через NuGet.");
        }

        public void MergePdfs(List<string> inputPaths, string outputPath)
        {
            using (var outputDocument = new PdfDocument())
            {
                foreach (var inputPath in inputPaths)
                {
                    using (var inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import))
                    {
                        for (int i = 0; i < inputDocument.PageCount; i++)
                        {
                            outputDocument.AddPage(inputDocument.Pages[i]);
                        }
                    }
                }

                outputDocument.Save(outputPath);
            }
        }

        public int GetPageCount(string pdfPath)
        {
            using (var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.ReadOnly))
            {
                return document.PageCount;
            }
        }

        public void CreateIdCopyPdf(Image frontSide, Image backSide, string outputPath)
        {
            using (var document = new PdfDocument())
            {
                var page = document.AddPage();

                // A4 в портретной ориентации
                page.Width = XUnit.FromMillimeter(210);
                page.Height = XUnit.FromMillimeter(297);

                using (var gfx = XGraphics.FromPdfPage(page))
                {
                    double pageWidth = page.Width.Point;
                    double pageHeight = page.Height.Point;
                    double margin = 20;
                    double spacing = 20;

                    // Доступная область для каждого изображения
                    double availableWidth = pageWidth - 2 * margin;
                    double availableHeight = (pageHeight - 2 * margin - spacing) / 2;

                    // Отрисовка лицевой стороны (верхняя половина)
                    DrawImageCentered(gfx, frontSide, margin, margin, availableWidth, availableHeight);

                    // Линия разделения
                    double middleY = margin + availableHeight + spacing / 2;
                    gfx.DrawLine(XPens.LightGray, margin, middleY, pageWidth - margin, middleY);

                    // Отрисовка оборотной стороны (нижняя половина)
                    DrawImageCentered(gfx, backSide, margin, margin + availableHeight + spacing, availableWidth, availableHeight);
                }

                document.Save(outputPath);
            }
        }

        private void DrawImageCentered(XGraphics gfx, Image image, double x, double y, double maxWidth, double maxHeight)
        {
            using (var ms = new MemoryStream())
            {
                image.Save(ms, ImageFormat.Png);
                ms.Position = 0;

                using (var xImage = XImage.FromStream(ms))
                {
                    double ratioX = maxWidth / xImage.PixelWidth;
                    double ratioY = maxHeight / xImage.PixelHeight;
                    double ratio = Math.Min(ratioX, ratioY);

                    double newWidth = xImage.PixelWidth * ratio;
                    double newHeight = xImage.PixelHeight * ratio;

                    double offsetX = x + (maxWidth - newWidth) / 2;
                    double offsetY = y + (maxHeight - newHeight) / 2;

                    gfx.DrawImage(xImage, offsetX, offsetY, newWidth, newHeight);
                }
            }
        }
    }
}
