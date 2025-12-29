using System.Windows.Media;
using MFPControlCenter.Services;

namespace MFPControlCenter.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private string _printerStatus = "Поиск принтера...";
        private Brush _printerStatusColor = Brushes.Orange;
        private string _statusMessage = "Готов к работе";
        private int _progress;
        private bool _isProgressVisible;

        public string PrinterStatus
        {
            get => _printerStatus;
            set => SetProperty(ref _printerStatus, value);
        }

        public Brush PrinterStatusColor
        {
            get => _printerStatusColor;
            set => SetProperty(ref _printerStatusColor, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public int Progress
        {
            get => _progress;
            set => SetProperty(ref _progress, value);
        }

        public bool IsProgressVisible
        {
            get => _isProgressVisible;
            set => SetProperty(ref _isProgressVisible, value);
        }

        public MainViewModel()
        {
            CheckPrinterStatus();
        }

        private void CheckPrinterStatus()
        {
            var printService = new PrintService();
            var printerName = printService.FindHPLaserJetPrinter();

            if (!string.IsNullOrEmpty(printerName))
            {
                PrinterStatus = "Подключён";
                PrinterStatusColor = Brushes.LimeGreen;
                StatusMessage = $"Принтер: {printerName}";
            }
            else
            {
                PrinterStatus = "Не найден";
                PrinterStatusColor = Brushes.Red;
                StatusMessage = "HP LaserJet M1536dnf не обнаружен";
            }
        }

        public void UpdateStatus(string message)
        {
            StatusMessage = message;
        }

        public void ShowProgress(int value)
        {
            Progress = value;
            IsProgressVisible = value > 0 && value < 100;
        }
    }
}
