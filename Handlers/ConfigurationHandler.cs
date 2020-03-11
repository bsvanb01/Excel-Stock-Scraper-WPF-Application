using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelStockScraper
{
    class ConfigurationHandler : ConfigurationSection
    {
        [ConfigurationProperty("tickers", Options = ConfigurationPropertyOptions.IsRequired)]
        public TickersCollection Tickers
        {
            get
            {
                return (TickersCollection)this["tickers"];
            }
        }

    }

    [ConfigurationCollection(typeof(TickerElement), AddItemName = "Ticker")]
    public class TickersCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new TickerElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }
            return ((TickerElement)element).Name;
        }
    }

    public class TickerElement : ConfigurationElement
    {
        [ConfigurationProperty("name", IsRequired = true, IsKey = true)]
        public string Name
        {
            get { return (string)base["name"]; }
        }

    }

}
