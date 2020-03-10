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

namespace ExcelStockScraper
{
    class MainExecutingClass : INotifyPropertyChanged
    {
        StockSiteScraperController control = new StockSiteScraperController();
        private ObservableCollection<StockData> _tickerCollection;
        public event PropertyChangedEventHandler PropertyChanged;
        private ICommand _addUserInputTicker;
        private ICommand _removeUserInputTicker;
        private string _loggingText;
        private string _userTextInput;
        private Thickness _itemMargins;

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

        public bool CanExecute
        {
            get
            {
                return true;
            }
        }

        #region Constructor
        public MainExecutingClass()
        {
            this.StockSiteScraperController = new StockSiteScraperController();
            TickerCollection = new ObservableCollection<StockData>();

            RunTaskASync(); 
        }
        #endregion

        public StockSiteScraperController StockSiteScraperController
        { get; set; }

        public void MainMethod()
        {
            try
            {
                //while(TickerCollection != null && TickerCollection.Count !=0)
                //{

                //        //Pull user input added to array
                //        TickerCollection = control.TickerCollection;

                //        control.UpdateTickerData();

                //        //Logs to logging textbox
                //        LoggingText = control.LoggingText();

                //        //updates values in excel
                //        //control.UpdateStockValue(TickerCollection);

                //        Thread.Sleep(10);

                //}
                while(true)
                {
                    TickerCollection = control.TickerCollection;
                    if (TickerCollection.Count > 0)
                    {

                        control.MainExecutingMethod();

                        LoggingText = control.LoggingText();
                        Thread.Sleep(10);
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
                                    control.PullTickerData(ticker);

                                    TickerCollection.Add(new StockData { Ticker = ticker, CurrentValue = control.CurrentValue });
                                }
                            }
                        }
                        else
                        {
                            control.PullTickerData(UserTextInput);

                            TickerCollection.Add(new StockData { Ticker = UserTextInput, CurrentValue = control.CurrentValue });
                        }
                    }
                }
            }
            catch(Exception ex)
            {

            }


            return TickerCollection;
        }
        public void AddToConfigSettings()
        {

        }

        public void RemoveFromConfigSettings()
        {

        }

        //public void CheckForConfigSettings()
        //{
        //    ConfigurationManager configSettings
        //}

        public ObservableCollection<StockData> RemoveTicker()
        {
            TickerCollection.Remove(new StockData { Ticker = "VOO" });
            return TickerCollection;
        }

        async Task RunTaskASync()
        {
            await Task.Run(() => MainMethod());
        }



    }
}
