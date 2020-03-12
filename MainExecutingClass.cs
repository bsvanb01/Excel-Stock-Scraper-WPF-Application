using System;
using ExcelStockScraper.Controllers;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Windows.Input;
using ExcelStockScraper.Handlers;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelStockScraper
{
    class MainExecutingClass : INotifyPropertyChanged
    {
        #region Properties
        StockSiteScraperController control = new StockSiteScraperController();
        private ObservableCollection<StockData> _tickerCollection;
        public event PropertyChangedEventHandler PropertyChanged;
        private ICommand _addUserInputTicker;
        private ICommand _removeUserInputTicker;
        private string _loggingText;
        private string _userTextInput;
        private int _activeColumn;
        private int _activeRow;
        private bool _isIntermediate;
        private StockData _selectedItemToRemove;

        private BackgroundWorker worker;

        public static Excel.Application oExcelApp;
        public static Excel.Workbook wb;
        public static Excel.Worksheet oSheet;

        #region Commands
        public ICommand AddUserInputTickerICommand
        {
            get
            {
                return _addUserInputTicker ?? (_addUserInputTicker = new CommandHandler(() => AddTickerCommand(), () => CanExecute));
            }
        }

        public ICommand RemoveTickerICommand
        {
            get
            {
                return _removeUserInputTicker ?? (_removeUserInputTicker = new CommandHandler(() => RemoveTickerCommand(), () => CanExecute));
            }
        }
        #endregion

        #region Collections
        public ObservableCollection<StockData> TickerCollection
        {
            get
            {
                return _tickerCollection;
            }
            set
            {
                _tickerCollection = value;
                OnPropertyChanged("TickerCollection");
            }
        }
        #endregion



        public StockData SelectedItemToRemove
        {
            get
            {
                return _selectedItemToRemove;
            }
            set
            {
                _selectedItemToRemove = value;
                OnPropertyChanged("SelectedItemToRemove");
            }
        }

        public string UserTextInput
        {
            get
            {
                return _userTextInput;
            }
            set
            {
                _userTextInput = value;
                OnPropertyChanged("UserTextInput");
            }

        }


        public string LoggingText
        {
            get
            {
                return _loggingText;
            }
            set
            {
                _loggingText = value;
                OnPropertyChanged("LoggingText");
            }
        }

        public int ActiveColumn
        {
            get
            {
                return _activeColumn;
            }
            set
            {
                _activeColumn = value;
                OnPropertyChanged("ActiveColumn");
            }
        }

        public int ActiveRow
        {
            get
            {
                return _activeRow;
            }
            set
            {
                _activeRow = value;
                OnPropertyChanged("ActiveRow");
            }
        }
        public bool IsIntermediate
        {
            get
            {
                return _isIntermediate;
            }
            set
            {
                _isIntermediate = value;
                OnPropertyChanged("IsIntermediate");
            }
        }

        public bool CanExecute
        {
            get
            {
                return true;
            }
        }

        #endregion

        #region Constructor
        public MainExecutingClass()
        {
            this.StockSiteScraperController = new StockSiteScraperController();
            TickerCollection = new ObservableCollection<StockData>();
            worker = CreateBackgroundWorker();
            worker.RunWorkerAsync();
            CheckForConfigSettings();
            RunTaskASync(); 
        }
        #endregion

        public StockSiteScraperController StockSiteScraperController
        { get; set; }


        private BackgroundWorker CreateBackgroundWorker()
        {
            var bw = new BackgroundWorker();
            bw.DoWork += worker_DoWork;
            bw.RunWorkerCompleted += worker_RunWorkerCompleted;
            return bw;
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            oExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            wb = oExcelApp.ActiveWorkbook;
            oSheet = oExcelApp.ActiveSheet;

            while (!worker.CancellationPending)
            {
                ActiveColumn = oSheet.Application.ActiveCell.Column;
                ActiveRow = oSheet.Application.ActiveCell.Row;
            }


        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ActiveColumn = ActiveColumn;
            ActiveRow = ActiveRow;
        }



        public void MainMethod()
        {
            try
            {
                while (true)
                {
                    TickerCollection = control.TickerCollection;
                    if (TickerCollection.Count >= 0)
                    {
                        control.UpdateTickerData();
                        
                        //LoggingText = control.LoggingText();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            

        }

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }


        public void CheckForConfigSettings()
        {
            try
            {
                var config = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location);
                var savedTickers = ConfigurationManager.GetSection("savedTickers") as ConfigurationHandler;
                var tickers = savedTickers.Tickers;

                if (tickers.Count != 0)
                {
                    IsIntermediate = true;
                    foreach (TickerElement key in tickers)
                    {
                        control.UserTickerInput.Add(key.Name);
                        control.AddTickersToCollection(key.Name.ToUpper());
                    }

                }
                else
                {

                }
            }
            catch(Exception ex)
            {
                LoggingText = ex.ToString();
            }

        }


        public ObservableCollection<StockData> AddTickerCommand()
        {
            
            try
            {
                if (!control.UserTickerInput.Contains(UserTextInput.ToUpper()))
                {
                    control.UserTickerInput.Add(UserTextInput.ToUpper());
                    if (control.UserTickerInput.Count != 0)
                    {
                        if (control.UserTickerInput.Count > 1)
                        {
                            foreach (string ticker in control.UserTickerInput)
                            {

                                if (!TickerCollection.Any(x => x.Ticker == ticker))
                                {
                                    control.AddTickersToCollection(UserTextInput.ToUpper());
                                }
                            }
                        }
                        else
                        {
                            control.AddTickersToCollection(UserTextInput.ToUpper());
                        }
                        control.AddToConfigSettings(UserTextInput.ToUpper());
                    }
                }
            }
            catch(Exception ex)
            {
                LoggingText = ex.ToString();
            }

            return TickerCollection;
        }

        public ObservableCollection<StockData> RemoveTickerCommand()
        {
            control.RemoveFromConfigSettings(SelectedItemToRemove);
            control.UserTickerInput.Remove(SelectedItemToRemove.Ticker);
            TickerCollection.Remove(SelectedItemToRemove);

            return TickerCollection;
        }

        async Task RunTaskASync()
        {
            await Task.Run(() => MainMethod());
        }




    }
}
