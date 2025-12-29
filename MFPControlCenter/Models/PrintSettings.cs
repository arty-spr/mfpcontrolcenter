using System.Collections.Generic;

namespace MFPControlCenter.Models
{
    public class PrintSettings
    {
        public string PrinterName { get; set; }
        public bool IsDuplex { get; set; }
        public int Copies { get; set; } = 1;
        public string PageRange { get; set; } = "all"; // "all", "1-5", "1,3,5-7"
        public PaperSize PaperSize { get; set; } = PaperSize.A4;
        public PrintQuality Quality { get; set; } = PrintQuality.Normal;
        public Orientation Orientation { get; set; } = Orientation.Portrait;
    }

    public enum PaperSize
    {
        A4,
        Letter,
        Legal,
        A5,
        B5
    }

    public enum PrintQuality
    {
        Draft,
        Normal,
        High
    }

    public enum Orientation
    {
        Portrait,
        Landscape
    }

    public static class PaperSizeInfo
    {
        public static Dictionary<PaperSize, string> Names = new Dictionary<PaperSize, string>
        {
            { PaperSize.A4, "A4 (210 x 297 мм)" },
            { PaperSize.Letter, "Letter (216 x 279 мм)" },
            { PaperSize.Legal, "Legal (216 x 356 мм)" },
            { PaperSize.A5, "A5 (148 x 210 мм)" },
            { PaperSize.B5, "B5 (176 x 250 мм)" }
        };
    }
}
