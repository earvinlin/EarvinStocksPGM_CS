using System;
using System.Collections.Generic;
using System.Text;
//using  EarvinStocksPGM.Modules;


namespace EarvinStocksPGM.Modules
{

    public static class GeneralModule
    {
        public const int MAP_K = 1;
        public const int MAP_VOLUME = 2;

        public class HighLowValues
        {
            public double highValue { get; set; }
            public double lowValue { get; set; }
        }


        public static HighLowValues GetHighLowValue(StockData[] sd, int startIndex, int displayCount, int type)
        {
            double highValue = 0;
            double lowValue = 99999;

            switch (type)
            {
                case MAP_K:
                    for (int i = startIndex; i < (startIndex + displayCount - 1); i++)
                    {
                        if (highValue < sd[i].HighPrice)
                            highValue = sd[i].HighPrice;
                        if (lowValue > sd[i].LowPrice)
                            lowValue = sd[i].LowPrice;
                    }
                    break;
                case MAP_VOLUME:
                    for (int i = startIndex; i < (startIndex + displayCount - 1); i++)
                    {
                        if (highValue < sd[i].Volume)
                            highValue = sd[i].Volume;
                        if (lowValue > sd[i].Volume)
                            lowValue = sd[i].Volume;
                    }
                    break;
            }
            HighLowValues values = new HighLowValues();
            values.highValue = highValue;
            values.lowValue = lowValue;

            return values;
        }

        public static int GetSelectFrame(FramePoints[] frmLeft, FramePoints[] frmRight, int frameNum)
        {
            int selectFrame = 0;


            return selectFrame;
        }

    }
}