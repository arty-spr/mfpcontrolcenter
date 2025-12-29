using System.Windows;
using MFPControlCenter.ViewModels;

namespace MFPControlCenter
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }
    }
}
