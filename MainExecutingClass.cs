using System;
using ExcelStockScraper.Controllers;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Windows.Input;
using ExcelStockScraper.Handlers;
using System.Windows.Data;
using System.Windows;
using System.Configuration;
using System.Xml;
using System.Collections.Specialized;

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
        private string _currentValue;
        private bool _isIntermediate;
        private StockData _selectedItemToRemove;


        public ICommand AddUserInputTickerCommand
        {
            get
            {
                return _addUserInputTicker ?? (_addUserInputTicker = new CommandHandler(() => AddTickerCommand(), () => CanExecute));
            }
        }

        public ICommand RemoveTickerCommand
        {
            get
            {
                return _removeUserInputTicker ?? (_removeUserInputTicker = new CommandHandler(() => RemoveTicker(), () => CanExecute));
            }
        }

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

        //public string CurrentValue

        //{
        //    get
        //    {
        //        return _currentValue;
        //    }
        //    set
        //    {
        //        _currentValue = value;
        //        OnPropertyChanged("CurrentValue");
        //    }
        //}

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
            CheckForConfigSettings();
            RunTaskASync(); 
        }
        #endregion

        public StockSiteScraperController StockSiteScraperController
        { get; set; }

        public void MainMethod()
        {
            try
            {
                while(true)
                {
                    TickerCollection = control.TickerCollection;
                    if (TickerCollection.Count > 0)
                    {
                        control.UpdateTickerData();
                        //foreach(StockData data in TickerCollection)
                        //{
                        //    CurrentValue = data.CurrentValue;
                        //}
                        //CurrentValue = control.CurrentValue;
                        LoggingText = control.LoggingText();
                        //Thread.Sleep(10);
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
            var config = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var savedTickers = ConfigurationManager.GetSection("savedTickers") as ConfigurationHandler;
            if (savedTickers == null)
            {

            }
            var tickers = savedTickers.Tickers;
            if (tickers.Count != 0)
            {
                IsIntermediate = true;
                foreach (TickerElement key in tickers)
                {
                    control.UserTickerInput.Add(key.Name);
                    control.AddTickersToCollection(key.Name);

                }
                
            }
            else
            {

            }

        }


        public ObservableCollection<StockData> AddTickerCommand()
        {
            
            try
            {
                if (!control.UserTickerInput.Contains(UserTextInput))
                {
                    control.UserTickerInput.Add(UserTextInput);
                    if (control.UserTickerInput.Count != 0)
                    {
                        if (control.UserTickerInput.Count > 1)
                        {
                            foreach (string ticker in control.UserTickerInput)
                            {

                                if (!TickerCollection.Any(x => x.Ticker == ticker))
                                {
                                    //TickerCollection.Add(new StockData { Ticker = UserTextInput, CurrentValue = control.PullTickerData(UserTextInput) });
                                    control.AddTickersToCollection(UserTextInput);
                                }
                            }
                        }
                        else
                        {
                            //control.PullTickerData(UserTextInput);

                            //TickerCollection.Add(new StockData { Ticker = UserTextInput, CurrentValue = control.PullTickerData(UserTextInput) });
                            control.AddTickersToCollection(UserTextInput);
                        }
                        control.AddToConfigSettings(UserTextInput);
                        
                    }
                }
            }
            catch(Exception ex)
            {

            }

            return TickerCollection;
        }




        public ObservableCollection<StockData> RemoveTicker()
        {
            control.RemoveFromConfigSettings(SelectedItemToRemove);
            TickerCollection.Remove(SelectedItemToRemove);
            
            
            return TickerCollection;
        }

        async Task RunTaskASync()
        {
            await Task.Run(() => MainMethod());
        }



    }
}
