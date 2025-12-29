using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Media.Imaging;

namespace MFPControlCenter.Helpers
{
    public static class ImageHelper
    {
        public static BitmapImage ConvertToBitmapImage(Image image)
        {
            if (image == null) return null;

            using (var ms = new MemoryStream())
            {
                image.Save(ms, ImageFormat.Png);
                ms.Position = 0;

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.StreamSource = ms;
                bitmapImage.EndInit();
                bitmapImage.Freeze();

                return bitmapImage;
            }
        }

        public static Image ConvertFromBitmapImage(BitmapImage bitmapImage)
        {
            if (bitmapImage == null) return null;

            using (var ms = new MemoryStream())
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmapImage));
                encoder.Save(ms);

                ms.Position = 0;
                return Image.FromStream(ms);
            }
        }

        public static Image RotateImage(Image image, float angle)
        {
            var bitmap = new Bitmap(image.Width, image.Height);

            using (var g = Graphics.FromImage(bitmap))
            {
                g.TranslateTransform(image.Width / 2f, image.Height / 2f);
                g.RotateTransform(angle);
                g.TranslateTransform(-image.Width / 2f, -image.Height / 2f);
                g.DrawImage(image, 0, 0);
            }

            return bitmap;
        }

        public static Image CropImage(Image image, Rectangle cropArea)
        {
            var bitmap = new Bitmap(cropArea.Width, cropArea.Height);

            using (var g = Graphics.FromImage(bitmap))
            {
                g.DrawImage(image, new Rectangle(0, 0, cropArea.Width, cropArea.Height),
                    cropArea, GraphicsUnit.Pixel);
            }

            return bitmap;
        }

        public static Image ResizeImage(Image image, int maxWidth, int maxHeight, bool preserveAspectRatio = true)
        {
            int newWidth, newHeight;

            if (preserveAspectRatio)
            {
                float ratioX = (float)maxWidth / image.Width;
                float ratioY = (float)maxHeight / image.Height;
                float ratio = Math.Min(ratioX, ratioY);

                newWidth = (int)(image.Width * ratio);
                newHeight = (int)(image.Height * ratio);
            }
            else
            {
                newWidth = maxWidth;
                newHeight = maxHeight;
            }

            var bitmap = new Bitmap(newWidth, newHeight);

            using (var g = Graphics.FromImage(bitmap))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(image, 0, 0, newWidth, newHeight);
            }

            return bitmap;
        }

        public static string GetFileSizeString(long bytes)
        {
            string[] sizes = { "Б", "КБ", "МБ", "ГБ" };
            int order = 0;
            double size = bytes;

            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }

            return $"{size:0.##} {sizes[order]}";
        }
    }
}
