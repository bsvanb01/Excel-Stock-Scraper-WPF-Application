using System;
using ExcelStockScraper.Controllers;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.ComponentModel;
using System.Collections.ObjectModel;

namespace ExcelStockScraper
{
    class MainExecutingClass : INotifyPropertyChanged
    {
        StockSiteScraperController control = new StockSiteScraperController();
        private ObservableCollection<StockData> _tickerCollection;
        public event PropertyChangedEventHandler PropertyChanged;
        private string _loggingText;

        public ObservableCollection<StockData> TickerCollection
        {
            get
            {
                return _tickerCollection;
            }
            set
            {
                _tickerCollection = value;
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

        public StockSiteScraperController StockSiteScraperController
        { get; set; }

        void brad()
        {
            
            while (true)
            {
                try
                {
                    TickerCollection =  control.StockDataCollection();
                    //control.ScrapeVOOFromWeb();
                    //control.ScrapeMGKFromWeb();
                    //control.ScrapeVONGFromWeb();
                    //control.ScrapeVUGFromWeb();
                    control.UpdateStockValue(StockSiteScraperController.VOO, StockSiteScraperController.MGK, StockSiteScraperController.VONG, StockSiteScraperController.VUG);
                    LoggingText = control.LoggingText();
                    Thread.Sleep(10);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }

            }

        }

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        async Task RunTaskASync()
        {
            await Task.Run(() => brad());
        }

        public MainExecutingClass()
        {
            
            RunTaskASync();

        }


    }
}
