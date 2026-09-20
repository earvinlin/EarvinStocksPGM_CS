using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace EarvinStocksPGM.Modules
{
    public static class GeneralModule
    {
        public struct FramePoints
        {
            public float frameX;
            public float frameY;
        }

        public const int MAP_UNSELECTED = 0;
        public const int MAP_KBAR = 1;
        public const int MAP_VOLUME = 2;
        public const int MAP_BIAS = 3;
        public const int MAP_WMS = 4;
        public const int MAP_PSY = 5;
        public const int MAP_SRSI = 6;
        public const int MAP_LRSI = 7;
        public const int MAP_RSI = 8;
        public const int MAP_K = 9;
        public const int MAP_D = 10;
        public const int MAP_KD = 11;

        // 記錄每個FRAME選擇顯示的資料(最多只能選9個；第1個一定是MAP_K)
        public static int[] SelectShowMapOnFrames = new int[9];

        public class HighLowValues
        {
            public double highValue { get; set; }
            public double lowValue { get; set; }
        }

        public static HighLowValues GetHighLowValue(StockData[] sd, IndexData[] idx, int startIndex, int displayCount, int type)
        {
            double highValue = 0;
            double lowValue = 99999;

            switch (type)
            {
                case MAP_KBAR :
                    for (int i = startIndex; i < (startIndex + displayCount - 1); i++)
                    {
                        if (highValue < sd[i].HighPrice)
                            highValue = sd[i].HighPrice;
                        if (lowValue > sd[i].LowPrice)
                            lowValue = sd[i].LowPrice;
                    }
                    break;
                case MAP_VOLUME :
                    for (int i = startIndex; i < (startIndex + displayCount - 1); i++)
                    {
                        if (highValue < sd[i].Volume)
                            highValue = sd[i].Volume;
                        if (lowValue > sd[i].Volume)
                            lowValue = sd[i].Volume;
                    }
                    break;
                case MAP_BIAS :
                    for (int i = startIndex; i < (startIndex + displayCount - 1); i++)
                    {
                        if (highValue < idx[i].BIAS)
                            highValue = idx[i].BIAS;
                        if (lowValue > idx[i].BIAS)
                            lowValue = idx[i].BIAS;
                    }
                    highValue = Math.Round(highValue, 2);
                    lowValue = Math.Round(lowValue, 2);

                    if (highValue >= 0 && lowValue >= 0)
                    {
                        highValue = Math.Ceiling(Math.Abs(highValue));
                        lowValue = -highValue;
                    } 
                    else if (highValue >= 0 && lowValue < 0)
                    {
                        if (Math.Abs(highValue) >= Math.Abs(lowValue))
                        {
                            highValue = Math.Ceiling(Math.Abs(highValue));
                            lowValue = -highValue;
                        }
                        else
                        {
                            highValue = Math.Ceiling(Math.Abs(lowValue));
                            lowValue = -highValue;
                        }
                    }
                    else 
                    {
                        //lowValue = Math.Floor(Math.Abs(lowValue));
                        //highValue = -lowValue;
                        highValue = Math.Ceiling(Math.Abs(lowValue));
                        lowValue = -highValue;
                    }
                    break;
                case MAP_WMS:
                    highValue = 100;
                    lowValue = 0;
                    break;
                case MAP_PSY:
                    highValue = 100;
                    lowValue = 0;
                    break;
                case MAP_SRSI:
                    highValue = 100;
                    lowValue = 0;
                    break;
                case MAP_LRSI:
                    highValue = 100;
                    lowValue = 0;
                    break;
                case MAP_K:
                    highValue = 100;
                    lowValue = 0;
                    break;
                case MAP_D:
                    highValue = 100;
                    lowValue = 0;
                    break;
                case MAP_KD:
                    highValue = 100;
                    lowValue = 0;
                    break;
            }
            HighLowValues values = new HighLowValues();
            values.highValue = highValue;
            values.lowValue = lowValue;

            return values;
        }

        public static int GetSelectFrame(FramePoints[] frmLeft, FramePoints[] frmRight, Point cursorPos, int frameNum)
        {
            int selectFrame = 0;

            Debug.WriteLine("GetSelectFrame()_INT--selectFrame= " + selectFrame);
            ////-- FOR DEBUG : Display Frame's 端點指標 --//
            //for (int i = 0; i < (frameNum + 1); i++)
            //{
            //    Debug.WriteLine("FramePoint[" + i + "], Left.X= " + frmLeft[i].frameX + ", Left.Y= " + frmLeft[i].frameY
            //                    + ", Right.X= " + frmRight[i].frameX + ", Right.Y= " + frmRight[i].frameY
            //                    + ", cursorPos.X= " + cursorPos.X + ", cursorPos.Y= " + cursorPos.Y);
            //}

            for (int i = 0; i < frameNum; i++)
            {
                if ((cursorPos.Y) < frmLeft[i].frameY)
                {
                    break;
                }
                else if ((cursorPos.Y) >= frmLeft[i].frameY && (cursorPos.Y) <= frmLeft[i + 1].frameY)
                {
                    selectFrame = (i + 1);
                    break;
                }
                else if ((cursorPos.Y) > frmLeft[frameNum].frameY)
                {
                    break;
                }
            }
            Debug.WriteLine("GetSelectFrame()_FIN--selectFrame= " + selectFrame);

            return selectFrame;
        }
    }
}