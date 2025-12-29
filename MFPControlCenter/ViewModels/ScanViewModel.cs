using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using MFPControlCenter.Helpers;
using MFPControlCenter.Models;
using MFPControlCenter.Services;
using Microsoft.Win32;
using ImageFormat = MFPControlCenter.Models.ImageFormat;

namespace MFPControlCenter.ViewModels
{
    public class ScanViewModel : BaseViewModel
    {
        private readonly ScanService _scanService;

        private ScanSource _selectedSource = ScanSource.Flatbed;
        private int _selectedDpi = 300;
        private ColorMode _selectedColorMode = ColorMode.Color;
        private ImageFormat _selectedFormat = ImageFormat.PDF;
        private BitmapImage _previewImage;
        private bool _isScanning;
        private string _statusMessage;
        private int _progress;
        private List<Image> _scannedPages = new List<Image>();
        private int _currentPageIndex;

        public ObservableCollection<ScanSource> Sources { get; }
        public ObservableCollection<int> DpiValues { get; }
        public ObservableCollection<ColorMode> ColorModes { get; }
        public ObservableCollection<ImageFormat> Formats { get; }

        public ICommand PreviewCommand { get; }
        public ICommand ScanCommand { get; }
        public ICommand ScanMultipleCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand ClearCommand { get; }
        public ICommand PreviousPageCommand { get; }
        public ICommand NextPageCommand { get; }

        public ScanSource SelectedSource
        {
            get => _selectedSource;
            set => SetProperty(ref _selectedSource, value);
        }

        public int SelectedDpi
        {
            get => _selectedDpi;
            set => SetProperty(ref _selectedDpi, value);
        }

        public ColorMode SelectedColorMode
        {
            get => _selectedColorMode;
            set => SetProperty(ref _selectedColorMode, value);
        }

        public ImageFormat SelectedFormat
        {
            get => _selectedFormat;
            set => SetProperty(ref _selectedFormat, value);
        }

        public BitmapImage PreviewImage
        {
            get => _previewImage;
            set => SetProperty(ref _previewImage, value);
        }

        public bool IsScanning
        {
            get => _isScanning;
            set => SetProperty(ref _isScanning, value);
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

        public int CurrentPageIndex
        {
            get => _currentPageIndex;
            set
            {
                if (SetProperty(ref _currentPageIndex, value))
                {
                    UpdatePreviewFromPages();
                }
            }
        }

        public int TotalPages => _scannedPages.Count;

        public string PageInfo => TotalPages > 0 ? $"Страница {CurrentPageIndex + 1} из {TotalPages}" : "Нет страниц";

        public ScanViewModel()
        {
            _scanService = new ScanService();
            _scanService.ScanProgress += OnScanProgress;

            Sources = new ObservableCollection<ScanSource>(Enum.GetValues(typeof(ScanSource)).Cast<ScanSource>());
            DpiValues = new ObservableCollection<int>(ScanSettingsInfo.AvailableDpi);
            ColorModes = new ObservableCollection<ColorMode>(Enum.GetValues(typeof(ColorMode)).Cast<ColorMode>());
            Formats = new ObservableCollection<ImageFormat>(Enum.GetValues(typeof(ImageFormat)).Cast<ImageFormat>());

            PreviewCommand = new AsyncRelayCommand(PreviewAsync, () => !IsScanning);
            ScanCommand = new AsyncRelayCommand(ScanAsync, () => !IsScanning);
            ScanMultipleCommand = new AsyncRelayCommand(ScanMultipleAsync, () => !IsScanning);
            SaveCommand = new RelayCommand(Save, () => _scannedPages.Count > 0);
            ClearCommand = new RelayCommand(Clear, () => _scannedPages.Count > 0);
            PreviousPageCommand = new RelayCommand(PreviousPage, () => CurrentPageIndex > 0);
            NextPageCommand = new RelayCommand(NextPage, () => CurrentPageIndex < TotalPages - 1);
        }

        private void OnScanProgress(object sender, ScanProgressEventArgs e)
        {
            Progress = e.Percent;
            StatusMessage = e.Message;
        }

        private async Task PreviewAsync()
        {
            IsScanning = true;
            StatusMessage = "Предпросмотр...";

            try
            {
                var settings = CreateSettings();
                var image = await Task.Run(() => _scanService.ScanPreview(settings));

                if (image != null)
                {
                    PreviewImage = ImageHelper.ConvertToBitmapImage(image);
                    StatusMessage = "Предпросмотр готов";
                    image.Dispose();
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
            }
            finally
            {
                IsScanning = false;
            }
        }

        private async Task ScanAsync()
        {
            IsScanning = true;
            StatusMessage = "Сканирование...";

            try
            {
                var settings = CreateSettings();
                var image = await Task.Run(() => _scanService.ScanSinglePage(settings));

                if (image != null)
                {
                    _scannedPages.Add(image);
                    CurrentPageIndex = _scannedPages.Count - 1;
                    OnPropertyChanged(nameof(TotalPages));
                    OnPropertyChanged(nameof(PageInfo));
                    StatusMessage = $"Отсканировано страниц: {TotalPages}";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
            }
            finally
            {
                IsScanning = false;
            }
        }

        private async Task ScanMultipleAsync()
        {
            IsScanning = true;
            StatusMessage = "Сканирование нескольких страниц...";

            try
            {
                var settings = CreateSettings();
                settings.Source = ScanSource.ADF; // Многостраничное только из ADF

                var images = await Task.Run(() => _scanService.ScanMultiplePages(settings));

                foreach (var image in images)
                {
                    _scannedPages.Add(image);
                }

                if (_scannedPages.Count > 0)
                {
                    CurrentPageIndex = _scannedPages.Count - 1;
                }

                OnPropertyChanged(nameof(TotalPages));
                OnPropertyChanged(nameof(PageInfo));
                StatusMessage = $"Отсканировано страниц: {TotalPages}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
            }
            finally
            {
                IsScanning = false;
            }
        }

        private void Save()
        {
            if (_scannedPages.Count == 0) return;

            var extension = GetExtension(SelectedFormat);
            var dialog = new SaveFileDialog
            {
                Title = "Сохранить скан",
                Filter = GetSaveFilter(),
                DefaultExt = extension,
                FileName = $"Scan_{DateTime.Now:yyyyMMdd_HHmmss}{extension}"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    if (_scannedPages.Count == 1)
                    {
                        _scanService.SaveScan(_scannedPages[0], dialog.FileName, SelectedFormat);
                    }
                    else
                    {
                        _scanService.SaveMultipleScans(_scannedPages, dialog.FileName, SelectedFormat);
                    }

                    StatusMessage = $"Сохранено: {dialog.FileName}";
                }
                catch (Exception ex)
                {
                    StatusMessage = $"Ошибка сохранения: {ex.Message}";
                }
            }
        }

        private void Clear()
        {
            foreach (var page in _scannedPages)
            {
                page.Dispose();
            }
            _scannedPages.Clear();
            PreviewImage = null;
            CurrentPageIndex = 0;
            OnPropertyChanged(nameof(TotalPages));
            OnPropertyChanged(nameof(PageInfo));
            StatusMessage = "Очищено";
        }

        private void PreviousPage()
        {
            if (CurrentPageIndex > 0)
            {
                CurrentPageIndex--;
            }
        }

        private void NextPage()
        {
            if (CurrentPageIndex < TotalPages - 1)
            {
                CurrentPageIndex++;
            }
        }

        private void UpdatePreviewFromPages()
        {
            if (_scannedPages.Count > 0 && CurrentPageIndex >= 0 && CurrentPageIndex < _scannedPages.Count)
            {
                PreviewImage = ImageHelper.ConvertToBitmapImage(_scannedPages[CurrentPageIndex]);
                OnPropertyChanged(nameof(PageInfo));
            }
        }

        private ScanSettings CreateSettings()
        {
            return new ScanSettings
            {
                Source = SelectedSource,
                Dpi = SelectedDpi,
                ColorMode = SelectedColorMode,
                Format = SelectedFormat
            };
        }

        private string GetExtension(ImageFormat format)
        {
            switch (format)
            {
                case ImageFormat.PDF: return ".pdf";
                case ImageFormat.JPEG: return ".jpg";
                case ImageFormat.PNG: return ".png";
                case ImageFormat.TIFF: return ".tiff";
                default: return ".pdf";
            }
        }

        private string GetSaveFilter()
        {
            return "PDF документ|*.pdf|JPEG изображение|*.jpg;*.jpeg|PNG изображение|*.png|TIFF изображение|*.tiff;*.tif";
        }
    }
}
