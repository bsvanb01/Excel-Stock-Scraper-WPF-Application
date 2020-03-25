using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Net;
using System.Text;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Configuration;
using System.Xml;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelStockScraper.Controllers
{
    public class ExcelData
    {
        public int SelectedColumn
        {
            get; set;
        }

        public int SelectedRow
        {
            get; set;
        }

    }

    public class StockData : INotifyPropertyChanged
    {

        private string _ticker;
        private string _currentValue;
        private string _gainLossValue;
        private string _gainLossValueColor;
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public string Ticker
        {
            get
            {
                return _ticker;
            }
            set
            {
                _ticker = value;
                OnPropertyChanged("Ticker");
            }
        }

        public string CurrentValue
        {
            get
            {
                return _currentValue;
            }
            set
            {
                _currentValue = value;
                OnPropertyChanged("CurrentValue");
            }
        }

        public string GainLossValue
        {
            get
            {
                return _gainLossValue;
            }
            set
            {
                _gainLossValue = value;
                OnPropertyChanged("GainLossValue");
            }
        }

        public string GainLossValueColor
        {
            get
            {
                return _gainLossValueColor;
            }
            set
            {
                _gainLossValueColor = value;
                OnPropertyChanged("GainLossValueColor");
            }
        }

        public string TickerExcelColumn
        {
            get; set;
        }

        public string TickerExcelRow
        {
            get; set;
        }

    }

    class StockSiteScraperController : INotifyPropertyChanged
    {

        #region Properties
        private static ObservableCollection<StockData> _tickerCollection;
        private static ObservableCollection<ExcelData> _excelDataCollection;
        
        private static List<string> _excelUpdateString;
        ObservableCollection<string> _userTickerInput;
        string loggingText = string.Empty;
        string _loggingTextString = string.Empty;
        
        private string stockValueElement = "//span[@class='Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)']";
        private string stockGainLossElement = "//span[starts-with(@class, 'Trsdu(0.3s) Fw(500) Pstart(10px) Fz(24px)')]";
        private int _loadingPercent;
        private int _activeColumn;
        private int _activeRow;
        int count = 1;

        XmlDocument _xmlDoc;
        HtmlAgilityPack.HtmlDocument doc;
        HtmlWeb web = new HtmlWeb();

        private BackgroundWorker worker;

        public static Excel.Application oExcelApp;
        public static Excel.Workbook wb;
        public static Excel.Worksheet oSheet;



        public event PropertyChangedEventHandler PropertyChanged;

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

        public ObservableCollection<ExcelData> ExcelDataCollection
        {
            get
            {
                return _excelDataCollection;
            }
            set
            {
                _excelDataCollection = value;
                OnPropertyChanged("ExcelDataCollection");
            }
        }

        public XmlDocument XmlDocument

        {
            get
            {
                return _xmlDoc;
            }
            set
            {
                _xmlDoc = value;
            }
        }

        public ObservableCollection<string> UserTickerInput
        {
            get
            {
                return _userTickerInput;
            }
            set
            {
                _userTickerInput = value;
                OnPropertyChanged("UserTickerInput");
            }
        }

        public List<string> ExcelUpdateString
        {
            get
            {
                return _excelUpdateString;
            }
            set
            {
                _excelUpdateString = value;
            }
        }


        public string LoggingTextString
        {
            get
            {
                return _loggingTextString;
            }
            set
            {
                _loggingTextString = value;
                OnPropertyChanged("LoggingTextString");
            }
        }



        public int LoadingPercent
        {
            get
            {
                return _loadingPercent;
            }
            set
            {
                _loadingPercent = value;
                OnPropertyChanged("LoadingPercent");
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

        public bool CanExecute
        {
            get
            {
                return true;
            }
        }

        #endregion

        #region Constructor
        public StockSiteScraperController()
        {
            XmlDocument = new XmlDocument();
            worker = CreateBackgroundWorker();
            worker.RunWorkerAsync();
            XmlDocument.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            TickerCollection = new ObservableCollection<StockData>();
            ExcelDataCollection = new ObservableCollection<ExcelData>();
            UserTickerInput = new ObservableCollection<string>();
            ExcelUpdateString = new List<string>();
        }
        #endregion

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }


        #region Methods

        public void UpdateTickerData()
        {
            for (int i = 0; i < TickerCollection.Count; i++)
            {
                //PullTickerData(TickerCollection[i].Ticker);
                if (TickerCollection.Any(x => x.Ticker == TickerCollection[i].Ticker))
                {
                    //TickerCollection[i].CurrentValue = 
                    PullTickerData(TickerCollection[i].Ticker, i);

                }
            }

        }


        public void PullTickerData(string ticker, int i)
        {

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                //doc = web.Load("https://finance.yahoo.com/quote/" + ticker + "/");

                var loadPage = Task.Run(() => web.Load("https://finance.yahoo.com/quote/" + ticker + "/"));
                doc = loadPage.Result;
                TickerCollection[i].CurrentValue = doc.DocumentNode.SelectSingleNode(stockValueElement).InnerHtml;
                TickerCollection[i].GainLossValue = doc.DocumentNode.SelectSingleNode(stockGainLossElement).InnerHtml;
                TickerCollection[i].GainLossValueColor = GainLossValueColorParse(TickerCollection[i].GainLossValue);
            }
            catch(Exception ex)
            {

            }

        }

        public string GainLossValueColorParse(string gainLossValue)
        {
            if(gainLossValue[0] == '-')
            {
                return "Red";
            }
            if(gainLossValue[0] == '+')
            {
                return "Green";
            }
            return "LightGray";
        }        


        public async Task AddTickersToCollection(string keyName, string tickerExcelRow, string tickerExcelColumn)
        {

            await Application.Current.Dispatcher.BeginInvoke(
                DispatcherPriority.Background, new Action(() =>
                {
                    TickerCollection.Add(new StockData { Ticker = keyName, TickerExcelRow = tickerExcelRow, TickerExcelColumn = tickerExcelColumn });
                }));
        }

        #region ConfigSettings
        //public void CheckForConfigSettings()
        //{
        //    var config = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location);
        //    var savedTickers = ConfigurationManager.GetSection("savedTickers") as ConfigurationHandler;
        //    if(savedTickers == null)
        //    {

        //    }
        //    var tickers = savedTickers.Tickers;
        //    if (tickers.Count != 0)
        //    {
        //        foreach (TickerElement key in tickers)
        //        {
        //            UserTickerInput.Add(key.Name);
        //            AddTickersToCollection(key.Name);
        //        }
        //    }
        //    else
        //    {

        //    }

        //}
        
        public void AddToConfigSettings(string ticker)
        {
            try
            {
                var nodeRegion = XmlDocument.CreateElement("Ticker");
                nodeRegion.SetAttribute("Name", ticker);
                nodeRegion.SetAttribute("ExcelRowValue", 0.ToString());
                nodeRegion.SetAttribute("ExcelColumnValue", 0.ToString());



                XmlDocument.SelectSingleNode("//savedTickers/tickers").AppendChild(nodeRegion);
                SaveAndRefresh("savedTickers/tickers");
            }
            catch (Exception ex)
            {
                LoggingTextString = ex.ToString();
            }

        }

        public void RemoveFromConfigSettings(StockData tickerName)
        {
            
            try
            {
                XmlNode nodeTicker = XmlDocument.SelectSingleNode("//savedTickers/tickers/Ticker[@Name=\'" + tickerName.Ticker + "\']");
                nodeTicker.ParentNode.RemoveChild(nodeTicker);
                UserTickerInput.Remove(tickerName.Ticker);
                SaveAndRefresh("savedTickers/tickers");

            }
            catch (Exception ex)
            {

                LoggingTextString = ex.ToString();
            }

            

        }

        public void SaveAndRefresh(string sectionToRefresh)
        {
            XmlDocument.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            ConfigurationManager.RefreshSection(sectionToRefresh);
        }

        #endregion

        #region Logger
        public string LoggingText()
        {
            
            if (count % 1 == 0)
            {
                if (TickerCollection.Count > 0)
                {
                    foreach (StockData stockData in TickerCollection)
                    {
                        LoggingTextString = LoggingTextString + " " + stockData.Ticker + ": " + stockData.CurrentValue;
                    }
                    LoggingTextString = LoggingTextString + "\n";

                    count++;
                }
                if (count == 100)
                {
                    count = 0;
                    
                }
                
            }
            return LoggingTextString;
        }
        #endregion

        #region Excel Updater
        public void UpdateStockValue(ObservableCollection<StockData> TickerCollection)
        {

            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();

            MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=K:\\BradFinances.xlsx; Extended Properties='Excel 12.0;HDR=YES; Mode=ReadWrite'");
            MyConnection.Open();
            myCommand.Connection = MyConnection;

            foreach(StockData stockData in TickerCollection)
            {
                ExcelUpdateString.Add("Update [Investment Data$] set Current_Stock_Prices = '" + stockData.CurrentValue + "' where Tickers = '" + stockData.Ticker + "'");
            }

            foreach (string str in ExcelUpdateString)
            {
                myCommand.CommandText = str;
                myCommand.ExecuteNonQuery();
            }

            MyConnection.Close();
        }
        #endregion

        #region Background Worker

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            oExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            wb = oExcelApp.ActiveWorkbook;
            oSheet = oExcelApp.ActiveSheet;
            
            while(!worker.CancellationPending)
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

        
        public void CheckForActiveCell()
        {
            while (worker.CancellationPending)
            {
                worker.RunWorkerAsync();
            }

            //var activeColumn = oSheet.Application.ActiveCell.Column;
            //var activeRow = oSheet.Application.ActiveCell.Row;

            //if (activeColumn != 0 && activeRow != 0)
            //{
            //    ActiveColumn = activeColumn;
            //    ActiveRow = activeRow;
            //}
        }

        private BackgroundWorker CreateBackgroundWorker()
        {
            var bw = new BackgroundWorker();
            bw.DoWork += worker_DoWork;
            bw.RunWorkerCompleted += worker_RunWorkerCompleted;
            return bw;
        }









        public void test()
        {
            oExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            wb = oExcelApp.ActiveWorkbook;
            oSheet = oExcelApp.ActiveSheet;

            var activeColumn = oSheet.Application.ActiveCell.Column;
            var activeRow = oSheet.Application.ActiveCell.Row;

            oSheet.Cells[activeRow, activeColumn] = "brad";
            
        }
        #endregion

        #endregion

    }
}
