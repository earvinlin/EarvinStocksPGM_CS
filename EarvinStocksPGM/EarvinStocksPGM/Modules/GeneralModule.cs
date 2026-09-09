using System;
using System.Collections.Generic;
using System.Text;
//using  EarvinStocksPGM.Modules;

namespace EarvinStocksPGM.Modules
{
    public static class GeneralModule
    {
        public class HighLowValues
        {
            public double highValue { get; set; }
            public double lowValue { get; set; }
        }


        public static HighLowValues GetHighLowValue(StockData[] sd, int startIndex, int displayCount, int type)
        {
            double stockPriceHighest = 0;
            double stockPriceLowest = 99999;

            // Call GetHighLow(sngLowPrice, sngHighPrice, dispMap, gudtStock, gudtIndex)
            for (int i = startIndex; i < (startIndex + displayCount - 1); i++)
            {
                if (stockPriceHighest < sd[i].HighPrice)
                    stockPriceHighest = sd[i].HighPrice;
                if (stockPriceLowest > sd[i].LowPrice)
                    stockPriceLowest = sd[i].LowPrice;
            }
            HighLowValues values = new HighLowValues();
            values.highValue = stockPriceHighest;
            values.lowValue = stockPriceLowest;
            return values;
        }
    }
}