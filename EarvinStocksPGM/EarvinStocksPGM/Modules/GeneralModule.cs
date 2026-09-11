using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        public static int GetSelectFrame(FramePoints[] frmLeft, FramePoints[] frmRight, Point cursorPos, int frameNum)
        {
            int selectFrame = 0;

            Debug.WriteLine("GGG_INT--selectFrame= " + selectFrame);
            //for (int i = 0; i < frameNum; i++)
            //{
            //    Debug.WriteLine("GGG--frameNum= " + frameNum + ", frmLeft[" + i + "].Y= " + frmLeft[i].frameY + ", frmLeft[" + (i+1) + "].Y= " + frmLeft[i + 1].frameY
            //        + ", cursorPos.X= " + cursorPos.X + ", cursorPos.Y= " + cursorPos.Y);
            //    if (cursorPos.Y < frmRight[i].frameY)
            //    {
            //        Debug.WriteLine("GGG0--frameNum= " + frameNum + ", frmLeft[" + i + "].Y= " + frmLeft[i].frameY + ", frmLeft[" + (i + 1) + "].Y= " + frmLeft[i + 1].frameY
            //            + ", cursorPos.X= " + cursorPos.X + ", cursorPos.Y= " + cursorPos.Y);
            //        break;
            //    }
            //    else if (cursorPos.Y >= frmRight[i].frameY && cursorPos.Y <= frmLeft[i + 1].frameY)
            //    {
            //        Debug.WriteLine("GGG1--frameNum= " + frameNum + ", frmLeft[" + i + "].Y= " + frmLeft[i].frameY + ", frmLeft[" + (i + 1) + "].Y= " + frmLeft[i + 1].frameY
            //            + ", cursorPos.X= " + cursorPos.X + ", cursorPos.Y= " + cursorPos.Y);
            //        selectFrame = (i+1);
            //        Debug.WriteLine("GGG_RETURN--selectFrame= " + selectFrame);
            //        break;
            //    }
            //    else if (cursorPos.Y > frmRight[frameNum].frameY)
            //    {
            //        Debug.WriteLine("GGG2--frameNum= " + frameNum + ", frmLeft[" + i + "].Y= " + frmLeft[i].frameY + ", frmLeft[" + (i + 1) + "].Y= " + frmLeft[i + 1].frameY
            //            + ", cursorPos.X= " + cursorPos.X + ", cursorPos.Y= " + cursorPos.Y);
            //        break;
            //    }
            //    Debug.WriteLine("GGG3--frmLeft[" + i + "].Y= " + frmLeft[i].frameY + ", frmLeft[" + (i + 1) + "].Y= " + frmLeft[i + 1].frameY
            //        + ", cursorPos.X= " + cursorPos.X + ", cursorPos.Y= " + cursorPos.Y);
            //}

            //-- FOR DEBUG : Display Frame's 端點指標 --//
            for (int i = 0; i < (frameNum + 1); i++)
            {
                Debug.WriteLine("FramePoint[" + i + "], Left.X= " + frmLeft[i].frameX + ", Left.Y= " + frmLeft[i].frameY
                                + ", Right.X= " + frmRight[i].frameX + ", Right.Y= " + frmRight[i].frameY
                                + ", cursorPos.X= " + cursorPos.X + ", cursorPos.Y= " + cursorPos.Y);
            }

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
            Debug.WriteLine("GGG_FIN--selectFrame= " + selectFrame);

            return selectFrame;
        }

    }
}