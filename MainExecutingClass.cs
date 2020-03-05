using System;
using ExcelStockScraper.Controllers;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.ComponentModel;

namespace ExcelStockScraper
{
    class MainExecutingClass
    {
        StockSiteScraperController control = new StockSiteScraperController();
        void brad()
        {
            int count = 0;

            while (true)
            {
                try
                {
                    control.ScrapeVOOFromWeb();
                    control.ScrapeMGKFromWeb();
                    control.ScrapeVONGFromWeb();
                    control.ScrapeVUGFromWeb();
                    control.UpdateStockValue(StockSiteScraperController.VOO, StockSiteScraperController.MGK, StockSiteScraperController.VONG, StockSiteScraperController.VUG);
                    Thread.Sleep(10);
                    count++;
                    if (count % 1 == 0)
                    {
                        control.LoggingText = count + "- VOO: " + StockSiteScraperController.VOO + " MGK: " + StockSiteScraperController.MGK + " VONG: " + StockSiteScraperController.VONG + " VUG: " + StockSiteScraperController.VUG;
                        //Console.WriteLine(count + "- VOO: " + StockSiteScraperController.VOO + " MGK: " + StockSiteScraperController.MGK + " VONG: " + StockSiteScraperController.VONG + " VUG: " + StockSiteScraperController.VUG);
                        if (count == 100)
                        {
                            count = 0;
                            Console.Clear();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }

            }

        }

        public MainExecutingClass()
        {
            brad();
        }


    }
}
