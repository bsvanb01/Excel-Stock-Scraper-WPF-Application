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

namespace ExcelStockScraper.Controllers
{

    public class StockData
    {

        public string Ticker
        {
            get; set;
        }

        public string CurrentValue
        {
            get; set;
        }

        public double ItemMarginTop
        {
            get; set;
        }

        public double ItemMarginBottom
        {
            get; set;
        }

        public Thickness ItemMargins
        {
            get; set;
        }

    }

    class StockSiteScraperController : INotifyPropertyChanged
    {

        #region Properties
        private static ObservableCollection<StockData> _tickerCollection;
        XmlDocument _xmlDoc;
        private static List<string> _excelUpdateString;
        List<string> _userTickerInput;
        
        private string _currentValue;
        private string _loggingTextString;
        private static string stockValueElement = "//span[@class='Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)']";
        int count = 1;
        HtmlAgilityPack.HtmlDocument doc;
        HtmlWeb web = new HtmlWeb();
        


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

        public List<string> UserTickerInput
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

        public bool CanExecute
        {
            get
            {
                return true;
            }
        }

        #endregion


        public StockSiteScraperController()
        {
            XmlDocument = new XmlDocument();
            XmlDocument.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            TickerCollection = new ObservableCollection<StockData>();
            UserTickerInput = new List<string>();
            ExcelUpdateString = new List<string>();
        }
        

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        
        #region Methods

        public void UpdateTickerData()
        {
            for (int i = 0; i < TickerCollection.Count;i++)
            {
                PullTickerData(TickerCollection[i].Ticker);
                if (TickerCollection.Any(x => x.Ticker == TickerCollection[i].Ticker))
                {
                    TickerCollection[i].CurrentValue = PullTickerData(TickerCollection[i].Ticker);
                }
            }

        }
        public string PullTickerData(string ticker)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            doc = web.Load("https://finance.yahoo.com/quote/" + ticker + "/");
            CurrentValue = doc.DocumentNode.SelectSingleNode(stockValueElement).InnerHtml;

            return CurrentValue;
        }


        public void CheckForConfigSettings()
        {
            var config = ConfigurationManager.OpenExeConfiguration(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var savedTickers = ConfigurationManager.GetSection("savedTickers") as ConfigurationHandler;
            if(savedTickers == null)
            {

            }
            var tickers = savedTickers.Tickers;
            if (tickers.Count != 0)
            {
                foreach (TickerElement key in tickers)
                {
                    UserTickerInput.Add(key.Name);
                    AddTickersToCollection(key.Name);
                }
            }
            else
            {

            }

        }

        private async Task AddTickersToCollection(string keyName)
        {
            
            await Application.Current.Dispatcher.BeginInvoke(
                DispatcherPriority.Background, new Action(() => {
                    TickerCollection.Add(new StockData { Ticker = keyName, CurrentValue = PullTickerData(keyName) });
                    }));
        }

        public void AddToConfigSettings(string ticker)
        {
            var nodeRegion = XmlDocument.CreateElement("Ticker");
            nodeRegion.SetAttribute("name", ticker);

            XmlDocument.SelectSingleNode("//savedTickers/tickers").AppendChild(nodeRegion);
            XmlDocument.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            ConfigurationManager.RefreshSection("savedTickers/tickers");
        }

        public void RemoveFromConfigSettings(StockData tickerName)
        {
            try
            {
                XmlNode nodeTicker = XmlDocument.SelectSingleNode("//savedTickers/tickers/Ticker[@name=\'" + tickerName.Ticker + "\']");
                nodeTicker.ParentNode.RemoveChild(nodeTicker);

                XmlDocument.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
                ConfigurationManager.RefreshSection("savedTickers/tickers");
            }
            catch(Exception ex)
            {

            }

        }
    

        #region Logger
        public string LoggingText()
        {
            string loggingText = string.Empty;
            if (count % 1 == 0)
            {
                if (TickerCollection.Count > 0)
                {
                    foreach (StockData stockData in TickerCollection)
                    {
                        loggingText = loggingText + " " + stockData.Ticker + ": " + stockData.CurrentValue;
                    }
                    loggingText = loggingText + "\n";

                    count++;
                }
                if (count == 100)
                {
                    count = 0;
                    
                }
                
            }
            return loggingText;
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

        #region delet
        public static void test()
        {
            Excel.Application oExcelApp = null;
            Excel.Workbook wb;
            oExcelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

            wb = oExcelApp.ActiveWorkbook;

            var workingSheet = wb.Sheets["Investment Data"];
            //Excel.Sheets
            Excel.Range range = workingSheet.UsedRange;

            var wbs = oExcelApp.Workbooks;


            Excel.Sheets s = wb.Worksheets;
            //Excel.Worksheet ws = (Excel.Worksheet)s.
            //oExcelApp = (Excel.Application)Activator.CreateInstance(type, true);
            //Excel.Range range = 
            oExcelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            oExcelApp = null;
        }
        #endregion

        #endregion

    }
}
