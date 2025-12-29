using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace MFPControlCenter.Views
{
    public partial class PrinterSettingsDialog : Window, INotifyPropertyChanged
    {
        private int _brightness;
        private int _contrast;
        private int _sharpness;

        public int Brightness
        {
            get => _brightness;
            set
            {
                if (_brightness != value)
                {
                    _brightness = value;
                    OnPropertyChanged();
                }
            }
        }

        public int Contrast
        {
            get => _contrast;
            set
            {
                if (_contrast != value)
                {
                    _contrast = value;
                    OnPropertyChanged();
                }
            }
        }

        public int Sharpness
        {
            get => _sharpness;
            set
            {
                if (_sharpness != value)
                {
                    _sharpness = value;
                    OnPropertyChanged();
                }
            }
        }

        public PrinterSettingsDialog()
        {
            InitializeComponent();
            DataContext = this;
        }

        public PrinterSettingsDialog(int brightness, int contrast, int sharpness) : this()
        {
            Brightness = brightness;
            Contrast = contrast;
            Sharpness = sharpness;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            Brightness = 0;
            Contrast = 0;
            Sharpness = 0;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
