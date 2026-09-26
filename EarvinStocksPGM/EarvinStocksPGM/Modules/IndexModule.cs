using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using static System.Net.Mime.MediaTypeNames;
using static EarvinStocksPGM.Modules.GeneralModule;


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
        public double PSY { get; set; }
        public double SRSI { get; set; }
        public double LRSI { get; set; }
        public double K { get; set; }
        public double D { get; set; }
        public double DIF { get; set; }
        public double MACD { get; set; }
        public double OSC { get; set; }
        public double SECTORS { get; set; }
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
            double[] dblPSY = new double[sd.Length];
            double[] dblSRSI = new double[sd.Length];
            double[] dblLRSI = new double[sd.Length];
            double[] dblK = new double[sd.Length];
            double[] dblD = new double[sd.Length];
            double[] dblDIF = new double[sd.Length];
            double[] dblMACD = new double[sd.Length]; 
            double[] dblOSC = new double[sd.Length];
            double[] dblSectors = new double[sd.Length];


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
            //dblBIAS = CalculateBIAS(sd, 10);
            //dblWMS = CalculateWMS(sd, 10);
            //dblPSY = CalculatePSY(sd, 10);
            //dblSRSI = CalculateRSI(sd, 6);
            //dblLRSI = CalculateRSI(sd, 20);
            //(dblK, dblD) = CalculateKD(sd, 9);
            //(dblDIF, dblMACD, dblOSC) = CalculateMACD(sd, 12, 26, 9);
            dblBIAS = CalculateBIAS(sd, BIASDay);
            dblWMS = CalculateWMS(sd, WMSDay);
            dblPSY = CalculatePSY(sd, PSYDay);
            dblSRSI = CalculateRSI(sd, SRSIDay);
            dblLRSI = CalculateRSI(sd, LRSIDay);
            (dblK, dblD) = CalculateKD(sd, KDay);
            (dblDIF, dblMACD, dblOSC) = CalculateMACD(sd, DIFDay, MACDDay, OSCDay);

            dblSectors = CalculateSectors(sd, idx);

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
                idx[i].PSY = dblPSY[i];
                idx[i].SRSI = dblSRSI[i];
                idx[i].LRSI = dblLRSI[i];
                idx[i].K = dblK[i];
                idx[i].D = dblD[i];
                idx[i].DIF = dblDIF[i];
                idx[i].MACD = dblMACD[i];
                idx[i].OSC = dblOSC[i];
                idx[i].SECTORS = dblSectors[i];
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
            //// DEBUG : Display the BIAS values for verification
            //for (i = 0; i < sd.Length; i++)
            //{
            //    Debug.WriteLine($"BIAS[{i}] = {Math.Round(dblValues[i], 2)}");
            //}
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

        public static double[] CalculatePSY(StockData[] sd, int intDayNo)
        {
            int i = 0;
            int upDays = 0;
            double dblAverage = 0;
            double[] dblValues = new double[sd.Length];

            while (i < sd.Length)
            {
                if (i < intDayNo)
                {
                    dblAverage = 50;
                }
                else
                {
                    upDays = 0;
                    for (int j = i; j > (i - intDayNo); j--)
                    {
                        //Debug.WriteLine("sd[" + j + "].Date= " + sd[j].TradeDate + ", i= " + sd[j].EndPrice + ", (i-1)= " + sd[j - 1].EndPrice + ",up= " + upDays);
                        if (sd[j].EndPrice > sd[j - 1].EndPrice)
                            upDays++;
                    }
                    //Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", upDays= " + upDays);
                    dblAverage = upDays / (double)intDayNo * 100;
                }
                dblValues[i] = dblAverage;

                i++;
            }
            //// DEBUG : Display the PSY values for verification
            //for (i = 0; i < sd.Length; i++)
            //{
            //    Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", PSY= " + Math.Round(dblValues[i], 2));
            //}
            return dblValues;
        }

        public static double[] CalculateRSI(StockData[] sd, int intDayNo)
        {
            int i = 0;
            double dblUpValue = 0, dblDownValue = 0, dblDiff = 0;
            double dblAverage;
            double[] dblValues = new double[sd.Length];

            Debug.WriteLine("intDayNo= " + intDayNo);
            while (i < sd.Length)
            {
                dblUpValue = 0;
                dblDownValue = 0;
                if (i < intDayNo)
                {
                    for (int j = 1; j < i; j++)
                    {
                        dblDiff = sd[j].EndPrice - sd[j - 1].EndPrice;
                        if (dblDiff > 0)
                            dblUpValue += dblDiff;
                        else
                            dblDownValue += Math.Abs(dblDiff);
                    }
                    dblUpValue = dblUpValue / i;
                    dblDownValue = dblDownValue / i;
               }
                else
                {
                    for (int j = i; j > (i - intDayNo); j--)
                    {
                        dblDiff = sd[j].EndPrice - sd[j - 1].EndPrice;
                        if (dblDiff > 0)
                            dblUpValue += dblDiff;
                        else
                            dblDownValue += Math.Abs(dblDiff);
                    }
                    dblUpValue = dblUpValue / intDayNo;
                    dblDownValue = dblDownValue / intDayNo;
                 }
                if (i != 0 && (dblUpValue + dblDownValue) != 0)
                    dblAverage = dblUpValue / (dblUpValue + dblDownValue) * 100;
                else
                    dblAverage = 50;

                dblValues[i] = dblAverage;

                i++;
            }
            //// DEBUG : Display the PSY values for verification
            //for (i = 0; i < sd.Length; i++)
            //{
            //    Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", RSI= " + Math.Round(dblValues[i], 2));
            //}
            return dblValues;
        }

        public static (double[] K, double[] D) CalculateKD(StockData[] sd, int period = 9)
        {
            double prevK = 50;
            double prevD = 50;

            double[] kValues = new double[sd.Length];
            double[] dValues = new double[sd.Length];

            for (int i = 0; i < sd.Length; i++)
            {
                // 前 N-1 天無法計算
                if (i < period - 1)
                {
                    //kValues[i] = double.NaN;
                    //dValues[i] = double.NaN;
                    kValues[i] = 0;
                    dValues[i] = 0;
                    continue;
                }

                double highest = double.MinValue;
                double lowest = double.MaxValue;

                for (int j = i - period + 1; j <= i; j++)
                {
                    highest = Math.Max(highest, sd[j].HighPrice);
                    lowest = Math.Min(lowest, sd[j].LowPrice);
                }

                double rsv;

                if (highest == lowest)
                {
                    rsv = 50;
                }
                else
                {
                    rsv = (sd[i].EndPrice - lowest) / (highest - lowest) * 100.0;
                }

                double k = prevK * 2.0 / 3.0 + rsv * 1.0 / 3.0;
                double d = prevD * 2.0 / 3.0 + k * 1.0 / 3.0;

                kValues[i] = k;
                dValues[i] = d;

                prevK = k;
                prevD = d;
            }
            //for (int i = 0; i < sd.Length; i++)
            //{
            //    Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", K= " + Math.Round(kValues[i], 2) + ", D= " + Math.Round(dValues[i], 2));
            //}
            return (kValues, dValues);
        }


        public static (double[] DIF, double[] MACD, double[] OSC) CalculateMACD(
        StockData[] sd, int emaShort = 12, int emaLong = 26,int emaSignal = 9)
        {
            if (sd == null)
                throw new ArgumentNullException(nameof(sd));

            if (sd.Length == 0)
                return (Array.Empty<double>(),
                        Array.Empty<double>(),
                        Array.Empty<double>());

            int count = sd.Length;

            double[] difValues = new double[count];
            double[] macdValues = new double[count];
            double[] oscValues = new double[count];

            // EMA初始值使用第一筆收盤價
            double emaS = sd[0].EndPrice;
            double emaL = sd[0].EndPrice;
            double macd = 0.0;

            difValues[0] = 0.0;
            macdValues[0] = 0.0;
            oscValues[0] = 0.0;

            double shortFactor = 2.0 / (emaShort + 1);
            double longFactor = 2.0 / (emaLong + 1);
            double signalFactor = 2.0 / (emaSignal + 1);

            for (int i = 1; i < count; i++)
            {
                double close = sd[i].EndPrice;

                // EMA Short
                emaS = emaS + (close - emaS) * shortFactor;

                // EMA Long
                emaL = emaL + (close - emaL) * longFactor;

                // DIF
                double dif = emaS - emaL;

                // MACD (Signal)
                macd = macd + (dif - macd) * signalFactor;

                // Histogram
                double osc = dif - macd;

                difValues[i] = dif;
                macdValues[i] = macd;     // MACD
                oscValues[i] = osc;   // OSC (MACD Bar)
            }
            //// DEBUG : Display the MACD values for verification
            //for (int i = 0; i < sd.Length; i++)
            //{
            //    Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate +
            //        ", dif= " + Math.Round(difValues[i], 5) +
            //        ", macd= " + Math.Round(macdValues[i], 5) +
            //        ", osc= " + Math.Round(oscValues[i], 5));
            //}
            return (difValues, macdValues, oscValues);
        }

        public static double[] CalculateSectors(StockData[] sd, IndexData[] idx)
        {
            int i = 0;
            double[] dblValues = new double[sd.Length];

            Debug.WriteLine("sd.Length= " + sd.Length);
            while (i < sd.Length)
            {
                if (i < 50)
                    dblValues[i] = 1.0;
                else if (i < 100)
                    dblValues[i] = -1.0;
                else
                    dblValues[i] = 1.0;

                i++;
            }
            // DEBUG : Display the PSY values for verification
            for (i = 0; i < sd.Length; i++)
            {
                Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", SECTORS= " + dblValues[i]);
            }
            return dblValues;
            /*
                         '--------------------------------------------------------------------
                        '* 記錄該點是屬於高點或低點
                        '--------------------------------------------------------------------
                        If (nowIndex1 > nowIndex2) Or _
                            ((nowIndex1 = nowIndex2) And (prevIndex1 < prevIndex2)) Then
                            signalFlag = HIGH_SIGNAL     ' 高點
                        Else
                            signalFlag = LOW_SIGNAL   ' 低點
                        End If
                        .sngSector = signalFlag

             */
        }






        //--- Write Here ---//




    }
}








//////// 唸書時的 MACD 版本 (未完成)
//////public static (double[] DIF, double[] MACD, double[] DIF_MACD) CalculateMACD(StockData[] sd,
//////    int period = 12, int ema_s = 12, int ema_l = 26, int ema = 5)
//////{
//////    double sngEMA_S;
//////    double sngEMA_L;
//////    double sngPreEMA_S;
//////    double sngPreEMA_L;
//////    double sngPreMACD;
//////    double sngMACD;
//////    double sngDIF;
//////    double sngDIF_MACD;
//////    double sngDI;
//////    int i;
//////    int j;

//////    double[] difValues = new double[sd.Length];
//////    double[] macdValues = new double[sd.Length];
//////    double[] dif_macdValues = new double[sd.Length];

//////    sngPreEMA_S = sd[0].EndPrice;
//////    sngPreEMA_L = sd[0].EndPrice;
//////    sngPreMACD = 0;
//////    j = 1;

//////    while (j <= sd.Length) 
//////    {
//////        sngDI = (sd[j].HighPrice + sd[j].LowPrice + sd[j].EndPrice * 2) / 4;
//////        sngEMA_S = sngPreEMA_S + (2 * (sngDI - sngPreEMA_S) / (1 + ema_s));
//////        sngEMA_L = sngPreEMA_L + (2 * (sngDI - sngPreEMA_L) / (1 + ema_l));
//////        sngDIF = sngEMA_S - sngEMA_L;
//////        //sngMACD = sngPreMACD + (2 * (sngDIF - sngPreMACD) / (1 + period));
//////        sngMACD = sngPreMACD + (2 * (sngDIF - sngPreMACD) / (1 + period));
//////        sngDIF_MACD = sngDIF - sngMACD;

//////        difValues[j] = sngDIF;
//////        macdValues[j] = sngMACD;
//////        dif_macdValues[j] = sngDIF_MACD;

//////        sngPreEMA_S = sngEMA_S;
//////        sngPreEMA_L = sngEMA_L;
//////        sngPreMACD = sngMACD;
//////        j++;
//////    }

//////    return (difValues, macdValues, dif_macdValues);
//////    /*

//////        sngPreEMA_S = udtStock(1).sngEndprice
//////        sngPreEMA_L = udtStock(1).sngEndprice
//////        sngPreMACD = 0
//////        j = 1
//////        While j <= intStockNo
//////            sngDI = (udtStock(j).sngHighPrice + udtStock(j).sngLowPrice + udtStock(j).sngEndprice * 2) / 4
//////            sngEMA_S = sngPreEMA_S + (2 * (sngDI - sngPreEMA_S) / (1 + intSEMANo))
//////            sngEMA_L = sngPreEMA_L + (2 * (sngDI - sngPreEMA_L) / (1 + intLEMANo))
//////            sngDIF = sngEMA_S - sngEMA_L
//////            sngMACD = sngPreMACD + (2 * (sngDIF - sngPreMACD) / (1 + intMACDNo))
//////            sngDIF_MACD = sngDIF - sngMACD

//////            udtIndex(j).sngDIF = sngDIF
//////            udtIndex(j).sngMACD = sngMACD
//////            udtIndex(j).sngDIF_MACD = sngDIF_MACD

//////            sngPreEMA_S = sngEMA_S
//////            sngPreEMA_L = sngEMA_L
//////            sngPreMACD = sngMACD
//////            j = j + 1
//////        Wend


//////     * 標準 MACD 的寫法應該類似
//////        for (int i = 1; i < sd.Length; i++)
//////        {
//////            double close = sd[i].EndPrice;

//////            ema12 = ema12 + (close - ema12) * 2.0 / 13.0;
//////            ema26 = ema26 + (close - ema26) * 2.0 / 27.0;

//////            dif = ema12 - ema26;

//////            dea = dea + (dif - dea) * 2.0 / 10.0;

//////            hist = dif - dea;

//////            difValues[i] = dif;
//////            macdValues[i] = dea;
//////            dif_macdValues[i] = hist;
//////        }
//////    *-------------------------------


//////     */
//////}






//////        public static (double[] K, double[] D) CalculateKD(StockData[] sd, int intDayNo)
//////        {
//////            int i = 0, j = 0;
//////            double dblPrevK = 50, dblPrevD = 50, dblRsv = 0;
//////            double dblK = 0, dblD = 0;
//////            double dblMax = 0, dblMin = 0;
//////            double[] dblKValues = new double[sd.Length];
//////            double[] dblDValues = new double[sd.Length];

//////            while (i < sd.Length)
//////            {
//////                dblMax = double.MinValue;
//////                dblMin = double.MaxValue;
//////                if (i < (intDayNo - 1))
//////                {
//////                    dblKValues[i] = double.NaN;
//////                    dblDValues[i] = double.NaN;
//////                    i++;
//////                    continue;
//////                }
//////                else
//////                {
////////                    for (j = i; j > (i - intDayNo); j--)
//////                    for (j = i - intDayNo + 1; j <= i; j++)
//////                    {
//////                        if (dblMax < sd[j].HighPrice)
//////                            dblMax = sd[j].HighPrice;
//////                        if (dblMin > sd[j].LowPrice)
//////                            dblMin = sd[j].LowPrice;
//////                    }
//////                }

//////                if (dblMax != dblMin)
//////                    dblRsv = (sd[i].EndPrice - dblMin) / (dblMax - dblMin) * 100;
//////                else
//////                    dblRsv = 50;

//////                Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", dblMax= " + Math.Round(dblMax, 2) +
//////                    ", dblMin= " + Math.Round(dblMin, 2) + ", EndPrice= " + sd[i].EndPrice + ", dblRsv= " + dblRsv);
//////                dblK = dblPrevK * 2 / 3 + dblRsv / 3;
//////                dblD = dblPrevD * 2 / 3 + dblK / 3;
//////                Debug.WriteLine("dblPrevK= " + dblPrevK + ", dblPrevD= " + dblPrevD + ", dblK= " + dblK + ", dblD= " + dblD);

//////                if (i < (intDayNo - 1))
//////                {
//////                    dblKValues[i] = double.NaN;
//////                    dblDValues[i] = double.NaN;
//////                } 
//////                else
//////                {
//////                    dblKValues[i] = dblK;
//////                    dblDValues[i] = dblD;
//////                }
//////                i++;
//////            }
//////            // DEBUG : Display the PSY values for verification
//////            for (i = 0; i < sd.Length; i++)
//////            {
//////                Debug.WriteLine("sd[" + i + "].Date= " + sd[i].TradeDate + ", K= " + Math.Round(dblKValues[i], 2) + ", D= " + Math.Round(dblDValues[i], 2));
//////            }

//////            return (dblKValues, dblDValues);
//////        }