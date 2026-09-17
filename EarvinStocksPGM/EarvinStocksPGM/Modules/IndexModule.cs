using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

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
        public double BIAS { get; set; }
        public double WMS { get; set; }
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
            double[] dblBIAS = new double[sd.Length];
            double[] dblWMS = new double[sd.Length];

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
            dblBIAS = CalculateBIAS(sd, 10);
            dblWMS = CalculateWMS(sd, 5);

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
                idx[i].BIAS = dblBIAS[i];
                idx[i].WMS = dblWMS[i];
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
            //// DEBUG : Display the dblAverage values for verification
            //for (i = 0; i < sd.Length; i++)
            //{
            //    Debug.WriteLine($"dblAverage[{i}] = {Math.Round(dblValues[i], 2)}");
            //}
            return dblValues;
        }

        public static double[] CalculateBIAS(StockData[] sd, int intDayNo)
        {
            int i = 0, j = 0;
            double dblAverage = 0;
            double[] dblValues = new double[sd.Length];

            while (i < sd.Length)
            {
                dblAverage = 0;
                if (i < intDayNo)
                {
                    for (j = 0; j <= i; j++)
                    {
                        dblAverage += sd[j].EndPrice;
                    }
                    dblAverage = Math.Round(dblAverage / (i + 1),2);
                }
                else
                {
                    for (j = i; j > (i - intDayNo); j--)
                    {
                        dblAverage += sd[j].EndPrice;
                    }
                    dblAverage = dblAverage / intDayNo;
                }
                dblValues[i] = Math.Round((sd[i].EndPrice - dblAverage) / dblAverage * 100,2);

                i++;
            }
            // DEBUG : Display the BIAS values for verification
            for (i = 0; i < sd.Length; i++)
            {
                Debug.WriteLine($"BIAS[{i}] = {Math.Round(dblValues[i], 2)}");
            }
            return dblValues;
        }

        public static double[] CalculateWMS(StockData[] sd, int intDayNo)
        {
            int i = 0;
            double dblMax = -1, dblMin = 999999;
            double dblAverage = 0;
            double[] dblValues = new double[sd.Length];

            while (i < sd.Length)
            {
                dblMax = -1;
                dblMin = 99999999;
                if (i < intDayNo) 
                {
                    for (int j = 0; j <= i; j++)
                    {
                        if (dblMax < sd[j].HighPrice)
                            dblMax = sd[j].HighPrice;
                        if (dblMin > sd[j].LowPrice)
                            dblMin = sd[j].LowPrice;
                    }
                }
                else
                {
                    for (int j = i; j > (i - intDayNo); j--)
                    {
                        if (dblMax < sd[j].HighPrice)
                            dblMax = sd[j].HighPrice;
                        if (dblMin > sd[j].LowPrice)
                            dblMin = sd[j].LowPrice;
                    }
                }
                if (dblMax != dblMin)
                {
                    //Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", EndPrice= " + sd[i].EndPrice  + 
                    //    ", dblMax = " + dblMax + ", dblMin= " + dblMin);
                    dblAverage = (dblMax - sd[i].EndPrice) / (dblMax - dblMin) * 100;
                }
                else
                {
                    dblAverage = 50;
                }
                dblValues[i] = 100 - dblAverage;

                i++;
            }
//            // DEBUG : Display the WMS values for verification
//            for (i = 0; i < sd.Length; i++)
//            {
////                Debug.WriteLine($"WMS[{i}] = {Math.Round(dblValues[i], 2)}");
//                Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", WMS= " + Math.Round(dblValues[i], 2));
//            }
            return dblValues;
        }

        //--- Write Here ---//

    }
}


/**
 最高/低價：29.7, 20.85
FramePoint[0], Left.X= 79, Left.Y= 79, Right.X= 79, Right.Y= 79, Mid.X= 79, Mid.Y= 79
FramePoint[1], Left.X= 415, Left.Y= 415, Right.X= 415, Right.Y= 415, Mid.X= 415, Mid.Y= 415
FramePoint[2], Left.X= 499, Left.Y= 499, Right.X= 499, Right.Y= 499, Mid.X= 499, Mid.Y= 499
FramePoint[3], Left.X= 583, Left.Y= 583, Right.X= 583, Right.Y= 583, Mid.X= 583, Mid.Y= 583
FramePoint[4], Left.X= 667, Left.Y= 667, Right.X= 667, Right.Y= 667, Mid.X= 667, Mid.Y= 667
FramePoint[5], Left.X= 751, Left.Y= 751, Right.X= 751, Right.Y= 751, Mid.X= 751, Mid.Y= 751

<1> (F00~F01)
<2> (F00~F01)
<3> (F00~F01)
<4> (F00~F01)
<5> (F00~F01)
 */