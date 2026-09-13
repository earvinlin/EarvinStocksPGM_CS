using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace EarvinStocksPGM.Modules
{

    public static class GeneralModule
    {
        public const int MAP_K = 1;
        public const int MAP_VOLUME = 2;
        public const int MAP_BIAS = 3;

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
                case MAP_K :
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
                    if (Math.Abs(highValue) >= Math.Abs(lowValue))
                    {
                        highValue = Math.Ceiling(Math.Abs(highValue));
                        lowValue = -highValue;
                    } else
                    {
                        lowValue = Math.Floor(Math.Abs(lowValue));
                        highValue = -lowValue;
                    }
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