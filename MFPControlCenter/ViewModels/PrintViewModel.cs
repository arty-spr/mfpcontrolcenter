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

namespace MFPControlCenter.ViewModels
{
    public class PrintViewModel : BaseViewModel
    {
        private readonly PrintService _printService;
        private readonly DocumentService _documentService;

        private string _selectedPrinter;
        private string _selectedFilePath;
        private string _fileName;
        private BitmapImage _previewImage;
        private bool _isDuplex;
        private int _copies = 1;
        private string _pageRange = "all";
        private PaperSize _selectedPaperSize = PaperSize.A4;
        private PrintQuality _selectedQuality = PrintQuality.Normal;
        private Orientation _selectedOrientation = Orientation.Portrait;
        private bool _isPrinting;
        private bool _isLoading;
        private string _statusMessage;
        private int _totalPages;
        private int _currentPageIndex;
        private List<Image> _pageImages = new List<Image>();

        public ObservableCollection<string> AvailablePrinters { get; }
        public ObservableCollection<PaperSize> PaperSizes { get; }
        public ObservableCollection<PrintQuality> Qualities { get; }
        public ObservableCollection<Orientation> Orientations { get; }

        public ICommand SelectFileCommand { get; }
        public ICommand PrintCommand { get; }
        public ICommand PreviousPageCommand { get; }
        public ICommand NextPageCommand { get; }
        public ICommand FirstPageCommand { get; }
        public ICommand LastPageCommand { get; }

        public string SelectedPrinter
        {
            get => _selectedPrinter;
            set => SetProperty(ref _selectedPrinter, value);
        }

        public string SelectedFilePath
        {
            get => _selectedFilePath;
            set
            {
                if (SetProperty(ref _selectedFilePath, value))
                {
                    FileName = Path.GetFileName(value);
                    _ = LoadPreviewAsync();
                }
            }
        }

        public string FileName
        {
            get => _fileName;
            set => SetProperty(ref _fileName, value);
        }

        public BitmapImage PreviewImage
        {
            get => _previewImage;
            set => SetProperty(ref _previewImage, value);
        }

        public bool IsDuplex
        {
            get => _isDuplex;
            set => SetProperty(ref _isDuplex, value);
        }

        public int Copies
        {
            get => _copies;
            set => SetProperty(ref _copies, Math.Max(1, value));
        }

        public string PageRange
        {
            get => _pageRange;
            set => SetProperty(ref _pageRange, value);
        }

        public PaperSize SelectedPaperSize
        {
            get => _selectedPaperSize;
            set => SetProperty(ref _selectedPaperSize, value);
        }

        public PrintQuality SelectedQuality
        {
            get => _selectedQuality;
            set => SetProperty(ref _selectedQuality, value);
        }

        public Orientation SelectedOrientation
        {
            get => _selectedOrientation;
            set => SetProperty(ref _selectedOrientation, value);
        }

        public bool IsPrinting
        {
            get => _isPrinting;
            set => SetProperty(ref _isPrinting, value);
        }

        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public int TotalPages
        {
            get => _totalPages;
            set
            {
                if (SetProperty(ref _totalPages, value))
                {
                    OnPropertyChanged(nameof(PageInfo));
                    OnPropertyChanged(nameof(HasMultiplePages));
                }
            }
        }

        public int CurrentPageIndex
        {
            get => _currentPageIndex;
            set
            {
                if (value >= 0 && value < TotalPages)
                {
                    if (SetProperty(ref _currentPageIndex, value))
                    {
                        UpdatePreviewImage();
                        OnPropertyChanged(nameof(PageInfo));
                        OnPropertyChanged(nameof(CurrentPageNumber));
                    }
                }
            }
        }

        public int CurrentPageNumber
        {
            get => _currentPageIndex + 1;
            set => CurrentPageIndex = value - 1;
        }

        public string PageInfo => TotalPages > 0 ? $"Страница {CurrentPageNumber} из {TotalPages}" : "";

        public bool HasMultiplePages => TotalPages > 1;

        public PrintViewModel()
        {
            _printService = new PrintService();
            _documentService = new DocumentService();

            AvailablePrinters = new ObservableCollection<string>();
            PaperSizes = new ObservableCollection<PaperSize>(Enum.GetValues(typeof(PaperSize)).Cast<PaperSize>());
            Qualities = new ObservableCollection<PrintQuality>(Enum.GetValues(typeof(PrintQuality)).Cast<PrintQuality>());
            Orientations = new ObservableCollection<Orientation>(Enum.GetValues(typeof(Orientation)).Cast<Orientation>());

            SelectFileCommand = new RelayCommand(SelectFile);
            PrintCommand = new AsyncRelayCommand(PrintAsync, () => CanPrint());
            PreviousPageCommand = new RelayCommand(PreviousPage, () => CurrentPageIndex > 0);
            NextPageCommand = new RelayCommand(NextPage, () => CurrentPageIndex < TotalPages - 1);
            FirstPageCommand = new RelayCommand(FirstPage, () => CurrentPageIndex > 0);
            LastPageCommand = new RelayCommand(LastPage, () => CurrentPageIndex < TotalPages - 1);

            LoadPrinters();
        }

        private void LoadPrinters()
        {
            AvailablePrinters.Clear();
            foreach (var printer in _printService.GetAvailablePrinters())
            {
                AvailablePrinters.Add(printer);
            }

            var hpPrinter = _printService.FindHPLaserJetPrinter();
            if (!string.IsNullOrEmpty(hpPrinter))
            {
                SelectedPrinter = hpPrinter;
            }
            else if (AvailablePrinters.Count > 0)
            {
                SelectedPrinter = AvailablePrinters[0];
            }
        }

        private void SelectFile()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Выберите файл для печати",
                Filter = "Все поддерживаемые|*.pdf;*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.tif;*.doc;*.docx;*.ppt;*.pptx;*.txt;*.rtf|" +
                         "Документы PDF|*.pdf|" +
                         "Изображения|*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.tif;*.gif|" +
                         "Word документы|*.doc;*.docx|" +
                         "PowerPoint презентации|*.ppt;*.pptx|" +
                         "Текстовые файлы|*.txt;*.rtf|" +
                         "Все файлы|*.*",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
            {
                SelectedFilePath = dialog.FileName;
            }
        }

        private async Task LoadPreviewAsync()
        {
            if (string.IsNullOrEmpty(SelectedFilePath) || !File.Exists(SelectedFilePath))
            {
                ClearPreview();
                return;
            }

            IsLoading = true;
            StatusMessage = "Загрузка документа...";

            // Очистка предыдущих страниц
            ClearPageImages();

            try
            {
                await Task.Run(() =>
                {
                    _pageImages = _documentService.GetPagePreviews(SelectedFilePath);
                });

                TotalPages = _pageImages.Count;
                CurrentPageIndex = 0;

                if (TotalPages > 0)
                {
                    UpdatePreviewImage();
                    StatusMessage = $"Загружено: {TotalPages} стр.";
                }
                else
                {
                    StatusMessage = "Не удалось загрузить документ";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка: {ex.Message}";
                ClearPreview();
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void UpdatePreviewImage()
        {
            if (_pageImages.Count > 0 && CurrentPageIndex >= 0 && CurrentPageIndex < _pageImages.Count)
            {
                PreviewImage = ImageHelper.ConvertToBitmapImage(_pageImages[CurrentPageIndex]);
            }
            else
            {
                PreviewImage = null;
            }
        }

        private void ClearPreview()
        {
            ClearPageImages();
            PreviewImage = null;
            TotalPages = 0;
            CurrentPageIndex = 0;
        }

        private void ClearPageImages()
        {
            foreach (var img in _pageImages)
            {
                img?.Dispose();
            }
            _pageImages.Clear();
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

        private void FirstPage()
        {
            CurrentPageIndex = 0;
        }

        private void LastPage()
        {
            CurrentPageIndex = TotalPages - 1;
        }

        private bool CanPrint()
        {
            return !IsPrinting &&
                   !IsLoading &&
                   !string.IsNullOrEmpty(SelectedFilePath) &&
                   File.Exists(SelectedFilePath) &&
                   !string.IsNullOrEmpty(SelectedPrinter);
        }

        private async Task PrintAsync()
        {
            IsPrinting = true;
            StatusMessage = "Печать...";

            try
            {
                var settings = new PrintSettings
                {
                    PrinterName = SelectedPrinter,
                    IsDuplex = IsDuplex,
                    Copies = Copies,
                    PageRange = PageRange,
                    PaperSize = SelectedPaperSize,
                    Quality = SelectedQuality,
                    Orientation = SelectedOrientation
                };

                await Task.Run(() => _printService.PrintFile(SelectedFilePath, settings));

                StatusMessage = "Печать завершена";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка печати: {ex.Message}";
            }
            finally
            {
                IsPrinting = false;
            }
        }
    }
}
