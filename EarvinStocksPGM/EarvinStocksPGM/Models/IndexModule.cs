using System;
using System.Collections.Generic;
using System.Text;

namespace EarvinStocksPGM.Models
{
    /**
     * IndexData index = new IndexData();
     * index.MAP5 = 23500.25f;
     */
    public class IndexData
    {
        public double MAP5 { get; set; }
        public double MAP10 { get; set; }
        public double MAP20 { get; set; }
        public double MAP60 { get; set; }
        public double MAP120 { get; set; }
        public double MAP240 { get; set; }
        public double MAV5 { get; set; }
        public double MAV10 { get; set; }
        public double MAV20 { get; set; }
        public double MAV60 { get; set; }
        public double MAV120 { get; set; }
    }

    public static class IndexModule
    {
        public static IndexData[] GetIndexData(StockData[] sd)
        {
            IndexData[] idx = new IndexData[sd.Length];

            return idx;
        }

        public static IndexData[] CalculateAverage()
        {
            return new IndexData[0];
        }
    }
}
