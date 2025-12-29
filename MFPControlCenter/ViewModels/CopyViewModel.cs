using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using MFPControlCenter.Helpers;
using MFPControlCenter.Models;
using MFPControlCenter.Services;

namespace MFPControlCenter.ViewModels
{
    public class CopyViewModel : BaseViewModel
    {
        private readonly CopyService _copyService;

        private CopyMode _selectedMode = CopyMode.Instant;
        private ScanSource _selectedSource = ScanSource.Flatbed;
        private int _copies = 1;
        private int _scalePercent = 100;
        private int _brightness = 0;
        private int _contrast = 0;
        private bool _isDuplex;
        private bool _isCopying;
        private string _statusMessage;
        private int _progress;

        public ObservableCollection<CopyMode> CopyModes { get; }
        public ObservableCollection<ScanSource> Sources { get; }

        public ICommand CopyCommand { get; }
        public ICommand IdCopyCommand { get; }

        public CopyMode SelectedMode
        {
            get => _selectedMode;
            set
            {
                if (SetProperty(ref _selectedMode, value))
                {
                    OnPropertyChanged(nameof(ModeDescription));
                }
            }
        }

        public string ModeDescription => CopySettingsInfo.GetModeDescription(SelectedMode);

        public ScanSource SelectedSource
        {
            get => _selectedSource;
            set => SetProperty(ref _selectedSource, value);
        }

        public int Copies
        {
            get => _copies;
            set => SetProperty(ref _copies, Math.Max(1, Math.Min(99, value)));
        }

        public int ScalePercent
        {
            get => _scalePercent;
            set => SetProperty(ref _scalePercent, Math.Max(25, Math.Min(400, value)));
        }

        public int Brightness
        {
            get => _brightness;
            set => SetProperty(ref _brightness, Math.Max(-50, Math.Min(50, value)));
        }

        public int Contrast
        {
            get => _contrast;
            set => SetProperty(ref _contrast, Math.Max(-50, Math.Min(50, value)));
        }

        public bool IsDuplex
        {
            get => _isDuplex;
            set => SetProperty(ref _isDuplex, value);
        }

        public bool IsCopying
        {
            get => _isCopying;
            set => SetProperty(ref _isCopying, value);
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

        public CopyViewModel()
        {
            _copyService = new CopyService();
            _copyService.CopyProgress += OnCopyProgress;

            CopyModes = new ObservableCollection<CopyMode>(Enum.GetValues(typeof(CopyMode)).Cast<CopyMode>());
            Sources = new ObservableCollection<ScanSource>(Enum.GetValues(typeof(ScanSource)).Cast<ScanSource>());

            CopyCommand = new AsyncRelayCommand(CopyAsync, () => !IsCopying);
            IdCopyCommand = new AsyncRelayCommand(IdCopyAsync, () => !IsCopying);
        }

        private void OnCopyProgress(object sender, CopyProgressEventArgs e)
        {
            Progress = e.Percent;
            StatusMessage = e.Message;
        }

        private async Task CopyAsync()
        {
            IsCopying = true;
            Progress = 0;

            try
            {
                var settings = CreateSettings();

                await Task.Run(() =>
                {
                    switch (SelectedMode)
                    {
                        case CopyMode.Instant:
                            _copyService.InstantCopy(settings);
                            break;
                        case CopyMode.Deferred:
                            _copyService.DeferredCopy(settings);
                            break;
                        case CopyMode.IdCopy:
                            _copyService.IdCopy(settings, ShowIdCopyPrompt);
                            break;
                    }
                });

                StatusMessage = "Копирование завершено";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
            }
            finally
            {
                IsCopying = false;
                Progress = 0;
            }
        }

        private async Task IdCopyAsync()
        {
            IsCopying = true;
            Progress = 0;

            try
            {
                var settings = CreateSettings();

                await Task.Run(() =>
                {
                    _copyService.IdCopy(settings, ShowIdCopyPrompt);
                });

                StatusMessage = "ID-копирование завершено";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
            }
            finally
            {
                IsCopying = false;
                Progress = 0;
            }
        }

        private void ShowIdCopyPrompt(string message)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                MessageBox.Show(message, "ID-копия", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }

        private CopySettings CreateSettings()
        {
            return new CopySettings
            {
                Mode = SelectedMode,
                Source = SelectedSource,
                Copies = Copies,
                ScalePercent = ScalePercent,
                Brightness = Brightness,
                Contrast = Contrast,
                IsDuplex = IsDuplex
            };
        }
    }
}
