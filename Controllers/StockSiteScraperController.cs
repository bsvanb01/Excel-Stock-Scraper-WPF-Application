using HtmlAgilityPack;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;

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

    }

    class StockSiteScraperController : INotifyPropertyChanged
    {

        #region Properties

        private static string _voo;
        private static string _mgk;
        private static string _vong;
        private static string _vug;
        private static ObservableCollection<StockData> _tickerCollection;
        string[] tickerArray;
        private static string stockValueElement = "//span[@class='Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)']";
        int count = 1;

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
            }
        }

        public static string VOO
        {
            get
            {
                return _voo;
            }
            set
            {
                _voo = value;
            }
        }
        public static string MGK
        {
            get
            {
                return _mgk;
            }
            set
            {
                _mgk = value;
            }
        }

        public static string VONG
        {
            get
            {
                return _vong;
            }
            set
            {
                _vong = value;
            }
        }
        public static string VUG
        {
            get
            {
                return _vug;
            }
            set
            {
                _vug = value;
            }
        }

        #endregion


        public StockSiteScraperController()
        {
            TickerCollection = new ObservableCollection<StockData>();
        }

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        
        #region Methods


        public ObservableCollection<StockData> StockDataCollection()
        {
            HtmlWeb web = new HtmlWeb();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            tickerArray = new string[] {"VOO","AAPL"};

            foreach(string ticker in tickerArray)
            {
                HtmlAgilityPack.HtmlDocument doc = web.Load("https://finance.yahoo.com/quote/" + ticker + "/");
                var currentValue = doc.DocumentNode.SelectSingleNode(stockValueElement).InnerHtml;
                TickerCollection.Add(new StockData { Ticker = ticker, CurrentValue = currentValue});
            }

            return TickerCollection;
        }

        public string ScrapeVOOFromWeb()
        {
            HtmlWeb web = new HtmlWeb();

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HtmlAgilityPack.HtmlDocument doc = web.Load("https://finance.yahoo.com/quote/voo/");
            VOO = doc.DocumentNode.SelectSingleNode(stockValueElement).InnerHtml;

            return VOO;
        }

        public string ScrapeMGKFromWeb()
        {
            HtmlWeb web = new HtmlWeb();

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HtmlAgilityPack.HtmlDocument doc = web.Load("https://finance.yahoo.com/quote/mgk/");
            MGK = doc.DocumentNode.SelectSingleNode(stockValueElement).InnerHtml;

            return MGK;
        }

        public string ScrapeVONGFromWeb()
        {
            HtmlWeb web = new HtmlWeb();

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HtmlAgilityPack.HtmlDocument doc = web.Load("https://finance.yahoo.com/quote/vong/");
            VONG = doc.DocumentNode.SelectSingleNode(stockValueElement).InnerHtml;

            return VONG;
        }

        public string ScrapeVUGFromWeb()
        {
            HtmlWeb web = new HtmlWeb();

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HtmlAgilityPack.HtmlDocument doc = web.Load("https://finance.yahoo.com/quote/vug/");
            VUG = doc.DocumentNode.SelectSingleNode(stockValueElement).InnerHtml;

            return VUG;
        }

        public string LoggingText()
        {
            string loggingText = string.Empty;
            if (count % 1 == 0)
            {
                //loggingText = count + "- VOO: " + VOO + " MGK: " + MGK + " VONG: " + VONG + " VUG: " + VUG;
                loggingText = "" + count;
                foreach(StockData stockData in TickerCollection)
                {
                    loggingText = loggingText + " " + count + " -" + stockData.Ticker + ": " + stockData.CurrentValue;
                }
                if (count == 100)
                {
                    count = 0;
                    Console.Clear();
                }
                count++;
            }
            return loggingText;
        }


        public void UpdateStockValue(string VOO, string MGK, string VONG, string VUG)
        {

            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();

            MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=K:\\BradFinances.xlsx; Extended Properties='Excel 12.0;HDR=YES; Mode=ReadWrite'");
            MyConnection.Open();
            myCommand.Connection = MyConnection;

            string[] updateStockValueArray = new string[4]
            {
                "Update [Investment Data$] set Current_Stock_Prices = '" + VOO + "' where Tickers = 'VOO'",
                "Update [Investment Data$] set Current_Stock_Prices = '" + MGK + "' where Tickers = 'MGK'",
                "Update [Investment Data$] set Current_Stock_Prices = '" + VONG + "' where Tickers = 'VONG'",
                "Update [Investment Data$] set Current_Stock_Prices = '" + VUG + "' where Tickers = 'VUG'"

            };


            foreach (string str in updateStockValueArray)
            {
                myCommand.CommandText = str;
                myCommand.ExecuteNonQuery();
            }

            MyConnection.Close();
        }


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

    }
}
