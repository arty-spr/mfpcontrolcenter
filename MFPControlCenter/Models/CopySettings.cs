namespace MFPControlCenter.Models
{
    public class CopySettings
    {
        public CopyMode Mode { get; set; } = CopyMode.Instant;
        public int Copies { get; set; } = 1;
        public int ScalePercent { get; set; } = 100; // 25-400%
        public int Brightness { get; set; } = 0; // -50 to +50
        public int Contrast { get; set; } = 0; // -50 to +50
        public bool IsDuplex { get; set; } = false;
        public ScanSource Source { get; set; } = ScanSource.Flatbed;
    }

    public enum CopyMode
    {
        Instant,    // Мгновенная копия (сканирует и сразу печатает)
        Deferred,   // Отложенная копия (сканирует всё, потом печатает)
        IdCopy      // ID-копия (2 стороны на 1 лист)
    }

    public static class CopySettingsInfo
    {
        public static string GetModeDisplayName(CopyMode mode)
        {
            switch (mode)
            {
                case CopyMode.Instant: return "Мгновенная копия";
                case CopyMode.Deferred: return "Отложенная копия";
                case CopyMode.IdCopy: return "ID-копия (2 стороны на 1 лист)";
                default: return mode.ToString();
            }
        }

        public static string GetModeDescription(CopyMode mode)
        {
            switch (mode)
            {
                case CopyMode.Instant: return "Сканирует страницу и сразу отправляет на печать";
                case CopyMode.Deferred: return "Сканирует все страницы, собирает в PDF, затем печатает";
                case CopyMode.IdCopy: return "Сканирует обе стороны документа и размещает на одном листе";
                default: return "";
            }
        }
    }
}
