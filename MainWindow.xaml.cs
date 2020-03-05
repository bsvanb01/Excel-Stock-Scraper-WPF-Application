using System.Windows;

namespace ExcelStockScraper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Show();
            DataContext = new MainExecutingClass();
            
        }
    }
}
