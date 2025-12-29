namespace MFPControlCenter.Models
{
    public class ScanSettings
    {
        public ScanSource Source { get; set; } = ScanSource.Flatbed;
        public int Dpi { get; set; } = 300;
        public ColorMode ColorMode { get; set; } = ColorMode.Color;
        public ImageFormat Format { get; set; } = ImageFormat.PDF;
        public string SavePath { get; set; }
        public bool ShowPreview { get; set; } = true;
    }

    public enum ScanSource
    {
        Flatbed,  // Планшет (стекло)
        ADF       // Автоподатчик документов
    }

    public enum ColorMode
    {
        Color,      // Цветной (48-bit)
        Grayscale,  // Градации серого (8-bit)
        BlackWhite  // Чёрно-белый (1-bit)
    }

    public enum ImageFormat
    {
        PDF,
        JPEG,
        PNG,
        TIFF
    }

    public static class ScanSettingsInfo
    {
        public static int[] AvailableDpi = { 75, 100, 150, 200, 300, 600, 1200 };

        public static string GetColorModeDisplayName(ColorMode mode)
        {
            switch (mode)
            {
                case ColorMode.Color: return "Цветной";
                case ColorMode.Grayscale: return "Градации серого";
                case ColorMode.BlackWhite: return "Чёрно-белый";
                default: return mode.ToString();
            }
        }

        public static string GetSourceDisplayName(ScanSource source)
        {
            switch (source)
            {
                case ScanSource.Flatbed: return "Планшет (стекло)";
                case ScanSource.ADF: return "Автоподатчик (ADF)";
                default: return source.ToString();
            }
        }
    }
}
