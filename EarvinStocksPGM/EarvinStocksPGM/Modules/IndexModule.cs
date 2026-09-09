using System;
using System.Collections.Generic;
using System.Text;

//namespace EarvinStocksPGM.Models
namespace EarvinStocksPGM.Modules
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
            double[] dblMAPValues5 = new double[sd.Length];
            double[] dblMAPValues10 = new double[sd.Length];
            double[] dblMAPValues20 = new double[sd.Length];
            double[] dblMAPValues60 = new double[sd.Length];
            double[] dblMAPValues120 = new double[sd.Length];
            double[] dblMAPValues240 = new double[sd.Length];
            double[] dblMAVValues5 = new double[sd.Length];
            double[] dblMAVValues10 = new double[sd.Length];
            double[] dblMAVValues20 = new double[sd.Length];
            double[] dblMAVValues60 = new double[sd.Length];
            double[] dblMAVValues120 = new double[sd.Length];

            dblMAPValues5 = CalculateAverage(sd, 5, true);
            dblMAPValues10 = CalculateAverage(sd, 10, true);
            dblMAPValues20 = CalculateAverage(sd, 20, true);
            dblMAPValues60 = CalculateAverage(sd, 60, true);
            dblMAPValues120 = CalculateAverage(sd, 120, true);
            dblMAPValues240 = CalculateAverage(sd, 240, true);
            dblMAVValues5 = CalculateAverage(sd, 5, false);
            dblMAVValues10 = CalculateAverage(sd, 10, false);
            dblMAVValues20 = CalculateAverage(sd, 20, false);
            dblMAVValues60 = CalculateAverage(sd, 60, false);
            dblMAVValues120 = CalculateAverage(sd, 120, false);

            for (int i = 0; i < sd.Length; i++)
            {
                idx[i] = new IndexData();
                idx[i].MAP5 = dblMAPValues5[i];
                idx[i].MAP10 = dblMAPValues10[i];
                idx[i].MAP20 = dblMAPValues20[i];
                idx[i].MAP60 = dblMAPValues60[i];
                idx[i].MAP120 = dblMAPValues120[i];
                idx[i].MAP240 = dblMAPValues240[i];
                idx[i].MAV5 = dblMAVValues5[i];
                idx[i].MAV10 = dblMAVValues10[i];
                idx[i].MAV20 = dblMAVValues20[i];
                idx[i].MAV60 = dblMAVValues60[i];
                idx[i].MAV120 = dblMAVValues120[i];
            }

            return idx;
        }
        
        /**
         * 計算平均值
         * @param sd 股票資料陣列
         * @param intDayNo 計算平均的天數
         * @param blnType true:計算收盤價平均 false:計算成交量平均
         * @return IndexData[] 平均值陣列
         */
        public static double[] CalculateAverage(StockData[] sd, int intDayNo, Boolean blnType)
        {
            int i = 0;
            double dblAverage = 0;
            double[] dblValues = new double[sd.Length];

            while (i < sd.Length)
            {
                dblAverage = 0;
                if (i < (intDayNo - 1))
                {
                    int j = 0;
                    for (j = 0; j <= i; j++)
                    {
                        if (blnType)
                            dblAverage += sd[j].EndPrice;
                        else
                            dblAverage += sd[j].Volume;
                    }
                    dblAverage = dblAverage / (i + 1);
                }
                else if (intDayNo != 0)
                {
                    for (int j = i; j > (i - intDayNo); j--)
                    {
                        if (blnType)
                            dblAverage += sd[j].EndPrice;
                        else
                            dblAverage += sd[j].Volume;
                    }
                    dblAverage = dblAverage / intDayNo;
                }
                else
                    dblAverage = 0;
                dblValues[i] = dblAverage;
                i++;
            }
            return dblValues;
        }
    }
}
