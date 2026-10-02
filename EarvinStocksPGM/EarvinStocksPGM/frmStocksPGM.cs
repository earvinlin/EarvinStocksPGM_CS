using EarvinStocksPGM.Modules;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using static DbHelper;
using static EarvinStocksPGM.Modules.GeneralModule;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using StockData = EarvinStocksPGM.Modules.StockData; // 指定 StockData 代表哪個類別

namespace EarvinStocksPGM
{

    public partial class frmStocksPGM : Form
    {
        //-- 動態元件 --//
        private Label lblStokInfo;              // 動態新增label元件：顯示股票資訊用
        //private Label lblHighPrice;             // 動態新增label元件：顯示股票最高價
        //private Label lblLowPrice;              // 動態新增label元件：顯示股票最低價

        private bool _initialized = false;

        private static int FrameNum = 5;            // 要顯示的frame數量
        private static int SelectFramePos = 0;      // 選擇的frame位置(1~FrameNum)
        private float XWidthBorder = 20;            // frame左、右兩邊預留的空間
        private float YHeightBorder = 30;           // frame最下面預留的空間
                                                    //        private float FrameXTop = 30;               // frame最左上角的X座標
        private float FrameXTop = 40;               // frame最左上角的X座標
        private float FrameRightBorder = 150;       // frame最左上角的Y座標
        private Boolean IsShowFocusLine = false;    // 是否顯示焦點線段
        int DisplayCount = 100;                     // 顯示的資料筆數
        int StartIndex = 0;                         // 顯示的資料起始索引
        private Point CursorPosition = new Point(); // 滑鼠游標位置
        float FrameTopXCoord = 0;                   // 視窗中Frame的最上方X座標
        float FrameTopYCoord = 0;                   // 視窗中Frame的最上方Y座標
        float FrameBottomXCoord = 0;                // 視窗中Frame的最下方X座標
        float FrameBottomYCoord = 0;                // 視窗中Frame的最下方Y座標
        float FrameXAxisWidth = 0;                  // Frame的X軸長度
        float FrameBarWidth = 0;                    // 儲存K-Bar的寬度

        FramePoints[] FrameLeftPoints = new FramePoints[FrameNum + 1];      // FramePoints結構陣列，存放frame左邊各個點的座標
        FramePoints[] FrameRightPoints = new FramePoints[FrameNum + 1];     // FramePoints結構陣列，存放frame右邊各個點的座標
        FramePoints[] FrameMiddlePoints = new FramePoints[FrameNum + 1];    // FramePoints結構陣列，存放frame中間各個點的座標

        StockData[] StkData;    // 要顯示的股票資料
        IndexData[] IdxData;    // 要顯示的指數資料

        public frmStocksPGM()
        {
            InitializeComponent();
            this.DoubleBuffered = true;

            this.StartPosition = FormStartPosition.CenterScreen;
            flpnlStocksBar.Width = this.Width;
            _initialized = true;
        }


        //--------------------------------------------------------------------------------------------------------------//
        // -- Form 內部使用的私有函數 (START) --------------------------------------------------------------------------//

        /**
         * Draw MAP均線
         * xPrev    :
         * xCurr    :
         * frameTopY:
         * yDistance:
         * highValue:
         * prevMAP  :
         * currMAP  :
         */
        private void DrawMAPLine(Graphics g, Pen pen, float xPrev, float xCurr,
            float frameTopY, float yDistance, double highValue, double prevMAP, double currMAP)
        {
            float yPrev = frameTopY + yDistance * (float)(highValue - prevMAP);
            float yCurr = frameTopY + yDistance * (float)(highValue - currMAP);

            g.DrawLine(pen, xPrev, yPrev, xCurr, yCurr);
        }

        // -- Form 內部使用的私有函數 ( END ) --------------------------------------------------------------------------//
        //--------------------------------------------------------------------------------------------------------------//


        private void cboSelectStock_SelectedIndexChanged(object sender, EventArgs e)
        {
            StartIndex = 0;
            // 取得要顯示的股票資料
            StkData = StockModule.GetStockData(cboSelectStock.Text);
            IdxData = IndexModule.GetIndexData(StkData);

            this.Invalidate();
        }

        private void btnFocus_Click(object sender, EventArgs e)
        {
            IsShowFocusLine = !IsShowFocusLine;
            if (!IsShowFocusLine)
                this.Invalidate();
        }

        private void frmStocksPGM_Load(object sender, EventArgs e)
        {
            // Setting Index's Days
            BIASDay = 10;
            WMSDay = 10;
            PSYDay = 5;
            SRSIDay = 6;
            LRSIDay = 20;
            KDay = 9;
            DDay = 9;
            DIFDay = 12;
            MACDDay = 26;
            OSCDay = 9;

            this.ContextMenuStrip = cntMenuStrip;

            flpnlStocksBar.Width = this.Width;
            FrameNum = int.Parse(cboFrameNum.Text);

            //// (20260929) 調整panel上面元件的高度 -- 不確定是否是比較好的做法…因為有時候panel上面的元件高度會跑掉
            //lblFrameNum.Top = (pnlStocksBar.ClientSize.Height - lblFrameNum.Height) / 2;
            //cboFrameNum.Top = (pnlStocksBar.ClientSize.Height - cboFrameNum.Height) / 2;
            //btnFocus.Top = (pnlStocksBar.ClientSize.Height - btnFocus.Height) / 2;
            //cboStocksFrom.Top = (pnlStocksBar.ClientSize.Height - cboStocksFrom.Height) / 2;
            //cboStocksType.Top = (pnlStocksBar.ClientSize.Height - cboStocksType.Height) / 2;
            //lblSelectStock.Top = (pnlStocksBar.ClientSize.Height - lblSelectStock.Height) / 2;
            //cboSelectStock.Top = (pnlStocksBar.ClientSize.Height - cboSelectStock.Height) / 2;
            //btnZoomIn.Top = (pnlStocksBar.ClientSize.Height - btnZoomIn.Height) / 2;
            //btnZoomOut.Top = (pnlStocksBar.ClientSize.Height - btnZoomOut.Height) / 2;
            //btnFore3.Top = (pnlStocksBar.ClientSize.Height - btnFore3.Height) / 2;
            //btnFore2.Top = (pnlStocksBar.ClientSize.Height - btnFore2.Height) / 2;
            //btnFore1.Top = (pnlStocksBar.ClientSize.Height - btnFore1.Height) / 2;
            //btnBack1.Top = (pnlStocksBar.ClientSize.Height - btnBack1.Height) / 2;
            //btnBack2.Top = (pnlStocksBar.ClientSize.Height - btnBack2.Height) / 2;
            //btnBack3.Top = (pnlStocksBar.ClientSize.Height - btnBack3.Height) / 2;

            // (20261001)
            //int xPos = lblFrameNum.Width +10;
            //cboFrameNum.Location = new Point(xPos, cboFrameNum.Location.Y);
            //xPos += cboFrameNum.Width + 10;
            //btnZoomIn.Location = new Point(xPos, btnZoomIn.Location.Y);

            /*
            lblFrameNum
            cboFrameNum
            btnZoomIn
            btnZoomOut
            btnFore3
            btnFore2
            btnFore1
            btnBack1
            btnBack2
            btnBack3
            btnFocus
            lblStrocksFrom
            cboStocksFrom
            cboStocksType
            lblSelectStock
            cboSelectStock
            */

            // 取得要顯示的股票資料
            StkData = StockModule.GetStockData(cboSelectStock.Text);
            IdxData = IndexModule.GetIndexData(StkData);

            // 新增顯示股票資訊的標籤
            lblStokInfo = new Label()
            {
                Name = "lblStokInfo",
                Text = "This is a test message!",
                AutoSize = true,
                Location = new Point(10, mnuStocksList.Size.Height + flpnlStocksBar.Size.Height)
            };
            this.Controls.Add(lblStokInfo);
        }

        private void frmStocksPGM_Resize(object sender, EventArgs e)
        {
            if (!_initialized)
                return;

            flpnlStocksBar.Width = this.Width;
            this.Invalidate();
        }

        private void frmStocksPGM_Paint(object sender, PaintEventArgs e)
        {
            Debug.WriteLine($"PAINT SelectFramePos={SelectFramePos}");

            Graphics g = e.Graphics;
            e.Graphics.Clear(this.BackColor);

            // lblStokInfo : 顯示股票資訊
            float frmYTop = mnuStocksList.Size.Height + flpnlStocksBar.Size.Height + lblStokInfo.Size.Height;
            // XWidthBorder : 表示frame左右皆各內縮 (XWidthBorder / 2) 個pixels
            float frmXWidth = this.ClientSize.Width - FrameXTop - (XWidthBorder / 2);
            // YHeightBorder : 表示frame最下面上調YHeightBorder個pixels
            float frmYHeight = this.ClientSize.Height - frmYTop - YHeightBorder;
            // Frame 數量
            FrameNum = int.Parse(cboFrameNum.Text);

            //---------------------//
            //-- 繪製 Frame 外框 --//
            //---------------------//
            FrameLeftPoints = new FramePoints[FrameNum + 1];  // FramePoints結構陣列，存放frame左邊各個點的座標
            FrameRightPoints = new FramePoints[FrameNum + 1]; // FramePoints結構陣列，存放frame右邊各個點的座標

            for (int i = 0; i < (FrameNum + 1); i++)
            {
                FrameLeftPoints[i].frameX = FrameXTop;
                if (i == 0)
                {
                    FrameLeftPoints[i].frameY = frmYTop;
                }
                else if (i == 1)
                {
                    FrameLeftPoints[i].frameY = frmYTop + frmYHeight / 2;
                }
                else
                {
                    FrameLeftPoints[i].frameY = frmYTop + (frmYHeight / 2) + (frmYHeight / 2) / (FrameNum - 1) * (i - 1);
                }
                // (For CHECK / DEBUG) 顯示Frame最左側端點座標
                g.FillEllipse(Brushes.BlueViolet, FrameLeftPoints[i].frameX, FrameLeftPoints[i].frameY, 5, 5);
            }

            for (int i = 0; i < (FrameNum + 1); i++)
            {
                FrameRightPoints[i].frameX = FrameXTop + frmXWidth;
                if (i == 0)
                {
                    FrameRightPoints[i].frameY = frmYTop;
                }
                else if (i == 1)
                {
                    FrameRightPoints[i].frameY = frmYTop + frmYHeight / 2;
                }
                else
                {
                    FrameRightPoints[i].frameY = frmYTop + (frmYHeight / 2) + (frmYHeight / 2) / (FrameNum - 1) * (i - 1);
                }
                // (For CHECK / DEBUG) 顯示Frame最右側端點座標
                g.FillEllipse(Brushes.BlueViolet, FrameRightPoints[i].frameX, FrameRightPoints[i].frameY, 5, 5);
            }

            // FramePoints結構陣列，存放frame中間各個點的座標(最後1個點是frame最右下角的點)
            FrameMiddlePoints = new FramePoints[FrameNum + 1];

            for (int i = 0; i < (FrameNum + 1); i++)
            {
                FrameMiddlePoints[i].frameX = FrameXTop + (frmXWidth - FrameRightBorder);
                if (i == 0)
                {
                    FrameMiddlePoints[i].frameY = frmYTop;
                }
                else if (i == 1)
                {
                    FrameMiddlePoints[i].frameY = frmYTop + frmYHeight / 2;
                }
                else if (i == FrameNum)
                {
                    FrameMiddlePoints[i].frameY = frmYTop + frmYHeight;
                }
                else
                {
                    FrameMiddlePoints[i].frameY = frmYTop + (frmYHeight / 2) + (frmYHeight / 2) / (FrameNum - 1) * (i - 1);
                }
                // (For CHECK / DEBUG) 顯示Frame內側端點座標
                g.FillEllipse(Brushes.BlueViolet, FrameMiddlePoints[i].frameX, FrameMiddlePoints[i].frameY, 5, 5);
            }
            // Frame外框
            g.DrawRectangle(Pens.Blue, FrameXTop, frmYTop, frmXWidth, frmYHeight);
            // FrameNum
            g.DrawLine(Pens.Magenta, FrameMiddlePoints[0].frameX, FrameMiddlePoints[0].frameY, FrameMiddlePoints[FrameNum].frameX, FrameMiddlePoints[FrameNum].frameY);
            for (int i = 1; i < FrameNum; i++)
            {
                g.DrawLine(Pens.Brown, FrameLeftPoints[i].frameX, FrameLeftPoints[i].frameY, FrameRightPoints[i].frameX, FrameRightPoints[i].frameY);
            }


            // 在第1次使用StartIndex前要確認其數值範圍在 0 ~ (資料總筆數 - 顯示在畫面的筆數)間
            if (StartIndex > (StkData.Length - DisplayCount))
            {
                StartIndex = StkData.Length - DisplayCount;
            }
            if (StartIndex < 0)
            {
                StartIndex = 0;
            }
            // 取得目前顯示區間最高價與最低價
            HighLowValues highLowValues = GeneralModule.GetHighLowValue(StkData, IdxData, StartIndex, DisplayCount, GeneralModule.MAP_KBAR);
            Debug.WriteLine($"最高/低價：{highLowValues.highValue}, {highLowValues.lowValue}");

            const int DashLineCount = 5;
            const float LabelOffsetY = 6f;
            const float LeftMargin = 10f;

            float topY = FrameLeftPoints[0].frameY;
            float bottomY = FrameLeftPoints[1].frameY;
            float frameHeight = bottomY - topY;

            float leftX = FrameLeftPoints[0].frameX;
            float rightX = FrameMiddlePoints[0].frameX;

            decimal maxPrice = (decimal)highLowValues.highValue;
            decimal minPrice = (decimal)highLowValues.lowValue;
            decimal priceStep = (maxPrice - minPrice) / DashLineCount;

            using Pen pen = new(Color.Black, 1)
            {
                DashStyle = DashStyle.Dash,
                DashPattern = new[] { 4f, 2f }
            };

            using System.Drawing.Font labelFont = new(this.Font.FontFamily, 6);
            // 畫最高價
            g.DrawString(maxPrice.ToString("F2"), labelFont, Brushes.Black, LeftMargin, topY);
            // 畫虛線及價格標示
            for (int i = 1; i < DashLineCount; i++)
            {
                float y = topY + (frameHeight * i / DashLineCount);

                g.DrawLine(pen, leftX, y, rightX, y);

                decimal price = maxPrice - priceStep * i;

                g.DrawString(price.ToString("F2"), labelFont, Brushes.Black, LeftMargin, y - LabelOffsetY);
            }
            // 畫最低價
            g.DrawString(minPrice.ToString("F2"), labelFont, Brushes.Black, LeftMargin, bottomY - 10);



            // 取得要顯示的股票資料
            if (StkData == null || StkData.Length == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }

            //--------------------------------------//
            //-- 顯示 K-Map (最上方 Frame) 柱狀圖 --//
            //--------------------------------------//
            float XAxisLength = FrameMiddlePoints[0].frameX - FrameLeftPoints[0].frameX;            // 顯示 K-Map's Frame 的X軸長度
            float YAxisLength = FrameLeftPoints[1].frameY - FrameLeftPoints[0].frameY;              // 顯示 K-Map's Frame 的Y軸長度
            float barWidth = XAxisLength / (float)DisplayCount;    // 要繪製K-Bar的寬度
            FrameBarWidth = barWidth;
            float barHeight = 0;                            // 要繪製K-Bar的高度
            float barXCoord = FrameLeftPoints[0].frameX;    // 要繪製K-Bar的X座標
            float barYCoord = 0;                            // 要繪製K-Bar的Y座標
            float yDistance = YAxisLength / (float)Math.Abs(highLowValues.highValue - highLowValues.lowValue); // 取得每個價格對應的Y軸距離

            for (int i = StartIndex; i < (StartIndex + DisplayCount); i++)
            {
                // X 座標
                if (i != StartIndex)
                {
                    barXCoord += barWidth;
                }
                // Y 座標
                if (StkData[i].StartPrice > StkData[i].EndPrice)
                    barYCoord = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - StkData[i].StartPrice));
                else
                    barYCoord = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - StkData[i].EndPrice));
                // 計算 K-Bar 的高度
                barHeight = yDistance * (float)Math.Abs(StkData[i].StartPrice - StkData[i].EndPrice);
                // 繪製 K-Bar
                if (barHeight != 0)
                {
                    if (StkData[i].StartPrice > StkData[i].EndPrice)
                    {
                        using Brush brush = new SolidBrush(Color.Green);
                        g.FillRectangle(brush, barXCoord, barYCoord, barWidth, barHeight);
                    }
                    else
                    {
                        using Brush brush = new SolidBrush(Color.Red);
                        g.FillRectangle(brush, barXCoord, barYCoord, barWidth, barHeight);
                    }
                }
                else
                {
                    g.DrawLine(Pens.Black, barXCoord, barYCoord, (barXCoord + barWidth), barYCoord);
                }
                // 繪製 K-Bar 最高價 to 最低價之線段
                float x0 = (barXCoord + barWidth / 2);
                //                float y0 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - StkData[i].HighPrice));
                float y0 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)(highLowValues.highValue - StkData[i].HighPrice));
                float x1 = (barXCoord + barWidth / 2);
                //float y1 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - StkData[i].LowPrice));
                float y1 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)(highLowValues.highValue - StkData[i].LowPrice));

                if (StkData[i].StartPrice > StkData[i].EndPrice)
                    g.DrawLine(Pens.Green, x0, y0, x1, y1);
                else
                    g.DrawLine(Pens.Red, x0, y0, x1, y1);

                // 顯示股價均線 (MAP) 的線段
                if (i > StartIndex)
                {
                    var prev = IdxData[i - 1];
                    var curr = IdxData[i];

                    float frameTopY = FrameLeftPoints[0].frameY;

                    float xPrev = barXCoord - barWidth / 2;
                    float xCurr = barXCoord + barWidth / 2;

                    using Pen penMAP5 = new Pen(Color.Blue, 2);
                    using Pen penMAP10 = new Pen(Color.Black, 2);
                    using Pen penMAP20 = new Pen(Color.Orange, 2);
                    using Pen penMAP60 = new Pen(Color.Green, 2);
                    using Pen penMAP120 = new Pen(Color.Brown, 2);
                    using Pen penMAP240 = new Pen(Color.Violet, 2);

                    DrawMAPLine(g, penMAP5, xPrev, xCurr, frameTopY, yDistance, highLowValues.highValue, prev.MAP5, curr.MAP5);
                    DrawMAPLine(g, penMAP10, xPrev, xCurr, frameTopY, yDistance, highLowValues.highValue, prev.MAP10, curr.MAP10);
                    DrawMAPLine(g, penMAP20, xPrev, xCurr, frameTopY, yDistance, highLowValues.highValue, prev.MAP20, curr.MAP20);
                    DrawMAPLine(g, penMAP60, xPrev, xCurr, frameTopY, yDistance, highLowValues.highValue, prev.MAP60, curr.MAP60);
                    DrawMAPLine(g, penMAP120, xPrev, xCurr, frameTopY, yDistance, highLowValues.highValue, prev.MAP120, curr.MAP120);
                    DrawMAPLine(g, penMAP240, xPrev, xCurr, frameTopY, yDistance, highLowValues.highValue, prev.MAP240, curr.MAP240);
                }
            }

            //-------------------------------------//
            //-- 顯示 K-Map (上方Frame) 的直線段 --//
            //-------------------------------------//
            int k = 0;
            long prevNum = 0, nextNum = 0;

            for (int i = StartIndex; i < (StartIndex + DisplayCount); i++)
            {
                if (i == StartIndex)
                {
                    // 取交易日期的最後兩碼，若為 20240101，則取 01
                    prevNum = StkData[i].TradeDate % 100;
                    nextNum = StkData[i].TradeDate % 100;
                    continue;
                }
                nextNum = StkData[i].TradeDate % 100;

                if (prevNum > nextNum)
                {
                    float x0 = FrameLeftPoints[0].frameX + (barWidth * (i - StartIndex));
                    float y0 = FrameLeftPoints[0].frameY;
                    float x1 = x0;
                    //                    float y1 = FrameLeftPoints[1].frameY;
                    float y1 = FrameLeftPoints[FrameNum].frameY;
                    //Pen pen = new Pen(Color.Black, 1);
                    g.DrawLine(pen, x0, y0, x1, y1);

                    // 顯示交易日期(年月)
                    string strnum = StkData[i].TradeDate.ToString();
                    g.DrawString(strnum.Substring(0, strnum.Length - 2), new System.Drawing.Font(this.Font.FontFamily, 6), Brushes.Black, (int)(x0 - (2 / 2)), (int)(y1 + 5));

                    k = k + 1;
                }
                prevNum = nextNum;
            }

            //=== 顯示 Frame START ===// 
            try
            {
                for (int i = 1; i <= FrameNum; i++)
                {
                    switch (GeneralModule.SelectShowMapOnFrames[i])
                    {
                        case GeneralModule.MAP_UNSELECTED:
                            break;
                        case GeneralModule.MAP_KBAR:
                            //Chalk_MAP_KBAR(e.Graphics, i, FrameNum);
                            break;
                        case GeneralModule.MAP_VOLUME:
                            Chalk_MAP_VOLUME(e.Graphics, i, FrameNum);
                            break;
                        case GeneralModule.MAP_BIAS:
                            //Chalk_MAP_BIAS(e.Graphics, i, FrameNum);
                            Chalk_MAP_CENTER_SINGLE_LINE(e.Graphics, i, FrameNum, GeneralModule.MAP_BIAS);
                            break;
                        case GeneralModule.MAP_WMS:
                            //Chalk_MAP_SINGLE_LINE(e.Graphics, i, FrameNum, GeneralModule.MAP_WMS);
                            Chalk_MAP_SINGLE_LINE_2(e.Graphics, i, FrameNum, GeneralModule.MAP_WMS);
                            break;
                        case GeneralModule.MAP_PSY:
                            //Chalk_MAP_SINGLE_LINE(e.Graphics, i, FrameNum, GeneralModule.MAP_PSY);
                            Chalk_MAP_SINGLE_LINE_2(e.Graphics, i, FrameNum, GeneralModule.MAP_PSY);
                            break;
                        case GeneralModule.MAP_SRSI:
                            //Chalk_MAP_SINGLE_LINE(e.Graphics, i, FrameNum, GeneralModule.MAP_SRSI);
                            Chalk_MAP_SINGLE_LINE_2(e.Graphics, i, FrameNum, GeneralModule.MAP_SRSI);
                            break;
                        case GeneralModule.MAP_RSI:
                            Chalk_MAP_DOUBLE_LINE(e.Graphics, i, FrameNum, GeneralModule.MAP_SRSI, GeneralModule.MAP_LRSI);
                            break;
                        case GeneralModule.MAP_K:
                            //Chalk_MAP_SINGLE_LINE(e.Graphics, i, FrameNum, GeneralModule.MAP_K);
                            Chalk_MAP_SINGLE_LINE_2(e.Graphics, i, FrameNum, GeneralModule.MAP_K);
                            break;
                        case GeneralModule.MAP_D:
                            //Chalk_MAP_SINGLE_LINE(e.Graphics, i, FrameNum, GeneralModule.MAP_D);
                            Chalk_MAP_SINGLE_LINE_2(e.Graphics, i, FrameNum, GeneralModule.MAP_D);
                            break;
                        case GeneralModule.MAP_KD:
                            Chalk_MAP_DOUBLE_LINE(e.Graphics, i, FrameNum, GeneralModule.MAP_K, GeneralModule.MAP_D);
                            break;
                        case GeneralModule.MAP_MACD:
                            Chalk_MAP_MACD_LINE(e.Graphics, i, FrameNum, GeneralModule.MAP_MACD);
                            break;
                        case GeneralModule.MAP_SECTORS:
                            Chalk_MAP_CENTER_BAR(e.Graphics, i, FrameNum, GeneralModule.MAP_SECTORS);
                            break;
                        case GeneralModule.MAP_MARGIN_PURCHASE:
                            Chalk_MAP_SINGLE_LINE_2(e.Graphics, i, FrameNum, GeneralModule.MAP_MARGIN_PURCHASE);
                            break;
                        case GeneralModule.MAP_SHORT_SELLING:
                            Chalk_MAP_SINGLE_LINE_2(e.Graphics, i, FrameNum, GeneralModule.MAP_SHORT_SELLING);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //-- (MouseMvoe Event) ----------------------------------------------------------------------------------------
            int curIndex = 0;
            if (IsShowFocusLine)
            {
                FrameTopXCoord = FrameLeftPoints[0].frameX;
                FrameTopYCoord = FrameLeftPoints[0].frameY;
                FrameBottomXCoord = FrameLeftPoints[FrameNum].frameX;
                FrameBottomYCoord = FrameLeftPoints[FrameNum].frameY;
                FrameXAxisWidth = XAxisLength;
                FrameBarWidth = barWidth;

                // 顯示滑鼠游標所在的K-Bar的索引
                if (CursorPosition.X <= (int)FrameTopXCoord)
                    CursorPosition.X = (int)FrameTopXCoord;
                if (CursorPosition.X >= (FrameTopXCoord + FrameXAxisWidth))
                    CursorPosition.X = (int)(FrameTopXCoord + FrameXAxisWidth) - 1;
                else if (CursorPosition.X > (FrameTopXCoord + FrameXAxisWidth))
                    CursorPosition.X = (int)(FrameTopXCoord + FrameXAxisWidth);
                curIndex = (int)((CursorPosition.X - FrameTopXCoord) / FrameBarWidth) + StartIndex;
                g.DrawLine(Pens.Brown, CursorPosition.X, FrameTopYCoord, CursorPosition.X, FrameBottomYCoord);
            }
            //-------------------------------------------------------------------------------------------------------------

            // Y-Length : 顯示frame的Y軸長度
            lblStokInfo.Text = "日期：" + StkData[curIndex].TradeDate + " 開 " + StkData[curIndex].StartPrice + " 高 " + StkData[curIndex].HighPrice + " 低 " + StkData[curIndex].LowPrice + " 收 " + StkData[curIndex].EndPrice;

            //---------------------------------------//
            //=== 最右側的各項指標數值顯示(START) ===//
            //---------------------------------------//
            lblMAP1.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1);
            lblMAP2.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height);
            lblMAP3.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height * 2);
            lblMAP4.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height * 3);
            lblMAP5.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height * 4);
            lblMAP6.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height * 5);
            lblMAP1.Text = "MAP  5: " + IdxData[curIndex].MAP5.ToString("F2");
            lblMAP2.Text = "MAP 10: " + IdxData[curIndex].MAP10.ToString("F2");
            lblMAP3.Text = "MAP 20: " + IdxData[curIndex].MAP20.ToString("F2");
            lblMAP4.Text = "MAP 60: " + IdxData[curIndex].MAP60.ToString("F2");
            lblMAP5.Text = "MAP120: " + IdxData[curIndex].MAP120.ToString("F2");
            lblMAP6.Text = "MAP240: " + IdxData[curIndex].MAP240.ToString("F2");

            lblMAV1.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height * 6);
            lblMAV2.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height * 7);
            lblMAV3.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height * 8);
            lblMAV4.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height * 9);
            lblMAV5.Location = new System.Drawing.Point((int)FrameMiddlePoints[0].frameX + 1, (int)FrameMiddlePoints[0].frameY + 1 + lblMAP1.Size.Height * 10);
            lblMAV1.Text = "MAV  5: " + IdxData[curIndex].MAV5.ToString("F2");
            lblMAV2.Text = "MAV 10: " + IdxData[curIndex].MAV10.ToString("F2");
            lblMAV3.Text = "MAV 20: " + IdxData[curIndex].MAV20.ToString("F2");
            lblMAV4.Text = "MAV 60: " + IdxData[curIndex].MAV60.ToString("F2");
            lblMAV5.Text = "MAV120: " + IdxData[curIndex].MAV120.ToString("F2");

            try
            {
                for (int i = 1; i <= FrameNum; i++)
                    {
                        switch (GeneralModule.SelectShowMapOnFrames[i])
                    {
                        case GeneralModule.MAP_UNSELECTED:
                            break;
                        case GeneralModule.MAP_K:
                            break;
                        case GeneralModule.MAP_VOLUME:
                            g.DrawString("VOL : " + $"{StkData[curIndex].Volume}", new System.Drawing.Font(this.Font.FontFamily, 6), Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            break;
                        case GeneralModule.MAP_BIAS:
                            g.DrawString("BIAS : " + $"{IdxData[curIndex].BIAS.ToString("F2")}", new System.Drawing.Font(this.Font.FontFamily, 6), Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            break;
                        case GeneralModule.MAP_WMS:
                            g.DrawString("WMS : " + $"{IdxData[curIndex].WMS.ToString("F")}", new System.Drawing.Font(this.Font.FontFamily, 6), Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            break;
                        case GeneralModule.MAP_PSY:
                            g.DrawString("PSY : " + $"{IdxData[curIndex].PSY.ToString("F0")}", new System.Drawing.Font(this.Font.FontFamily, 6), Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            break;
                        case GeneralModule.MAP_RSI:
                            System.Drawing.Font f = new System.Drawing.Font(this.Font.FontFamily, 6);
                            int lineHeight = f.Height;
                            g.DrawString($"SRSI : {IdxData[curIndex].SRSI:F2}", f, Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            g.DrawString($"LRSI : {IdxData[curIndex].LRSI:F2}", f, Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2 + lineHeight);
                            break;
                        case GeneralModule.MAP_KD:
                            f = new System.Drawing.Font(this.Font.FontFamily, 6);
                            lineHeight = f.Height;
                            g.DrawString($"K : {IdxData[curIndex].K:F2}", f, Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            g.DrawString($"D : {IdxData[curIndex].D:F2}", f, Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2 + lineHeight);
                            break;
                        case GeneralModule.MAP_MACD:
                            f = new System.Drawing.Font(this.Font.FontFamily, 6);
                            lineHeight = f.Height;
                            g.DrawString($"DIF : {IdxData[curIndex].DIF:F2}", f, Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            g.DrawString($"MACD : {IdxData[curIndex].MACD:F2}", f, Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2 + lineHeight);
                            g.DrawString($"OSC : {IdxData[curIndex].OSC:F2}", f, Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2 + lineHeight * 2);
                            break;
                        case GeneralModule.MAP_SECTORS:
                            g.DrawString("SECTORS : " + $"{IdxData[curIndex].SECTORS.ToString("F0")}", new System.Drawing.Font(this.Font.FontFamily, 6), Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            break;
                        case GeneralModule.MAP_MARGIN_PURCHASE:
                            g.DrawString("MARGIN_PURCHASE : \n" + $"{StkData[curIndex].MarginPurchase.ToString("F0")}", new System.Drawing.Font(this.Font.FontFamily, 6), Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            break;
                        case GeneralModule.MAP_SHORT_SELLING:
                            g.DrawString("SHORT_SELLING : \n" + $"{StkData[curIndex].ShortSelling.ToString("F0")}", new System.Drawing.Font(this.Font.FontFamily, 6), Brushes.Black, FrameMiddlePoints[i - 1].frameX + 2, FrameMiddlePoints[i - 1].frameY + 2);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void cboFrameNum_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Invalidate();
        }

        private void frmStocksPGM_MouseMove(object sender, MouseEventArgs e)
        {
            CursorPosition = e.Location;
            // 滑鼠移動時，顯示焦點線段
            if (IsShowFocusLine)
                this.Invalidate();
        }

        private void btnBack1_Click(object sender, EventArgs e)
        {
            StartIndex += 1;
            this.Invalidate();
        }

        private void btnFore1_Click(object sender, EventArgs e)
        {
            StartIndex -= 1;
            this.Invalidate();
        }

        private void btnBack2_Click(object sender, EventArgs e)
        {
            StartIndex += 5;
            this.Invalidate();
        }

        private void btnFore2_Click(object sender, EventArgs e)
        {
            StartIndex -= 5;
            this.Invalidate();
        }

        private void btnBack3_Click(object sender, EventArgs e)
        {
            StartIndex = int.MaxValue;
            this.Invalidate();
        }

        private void btnFore3_Click(object sender, EventArgs e)
        {
            StartIndex = 0;
            this.Invalidate();
        }

        private void btnZoomOut_Click(object sender, EventArgs e)
        {
            DisplayCount = DisplayCount + DisplayCount / 10;
            this.Invalidate();
        }

        private void btnZoomIn_Click(object sender, EventArgs e)
        {
            DisplayCount = DisplayCount - DisplayCount / 10;
            if (DisplayCount < 10)
                DisplayCount = 10;
            this.Invalidate();
        }

        private void VolumeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_VOLUME;
            Debug.WriteLine($"CLICK VolumeToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void BIASToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_BIAS;
            Debug.WriteLine($"CLICK BIASToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void WMSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_WMS;
            Debug.WriteLine($"CLICK BIASToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void PSYToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_PSY;
            Debug.WriteLine($"CLICK PSYToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void RSIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_RSI;
            Debug.WriteLine($"CLICK RSIToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void KDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_KD;
            Debug.WriteLine($"CLICK KDToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void MACDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_MACD;
            Debug.WriteLine($"CLICK MACDToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void SectorsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_SECTORS;
            Debug.WriteLine($"CLICK SectorsToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void MarginPurchaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_MARGIN_PURCHASE;
            Debug.WriteLine($"CLICK MarginPurchaseToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void ShortSellingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            SelectShowMapOnFrames[SelectFramePos] = GeneralModule.MAP_SHORT_SELLING;
            Debug.WriteLine($"CLICK ShortSellingToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void cboStocksType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cboStocksFrom_SelectedIndexChanged(object sender, EventArgs e)
        {

        }











        //========================================================================================================================
        //========================================================================================================================

        // 繪製成交量柱狀圖 (MAP_VOLUME)
        private void Chalk_MAP_VOLUME(Graphics g, int framePos, int frameNum)
        {
            // 邊界與參數安全檢查
            if (framePos <= 0 || framePos > frameNum || DisplayCount <= 0)
                return;

            // Y 軸高度 (注意：GDI+ 座標系頂端 Y 較小，底端 Y 較大)
            float frameTopY = FrameLeftPoints[framePos - 1].frameY;
            float frameBottomY = FrameLeftPoints[framePos].frameY;
            float yAxisLength = frameBottomY - frameTopY;

            // 取得最高/最低成交量
            HighLowValues highLowValues = GeneralModule.GetHighLowValue(StkData, IdxData, StartIndex, DisplayCount, GeneralModule.MAP_VOLUME);
            double maxVol = highLowValues.highValue;
            double minVol = highLowValues.lowValue;
            double volRange = Math.Abs(maxVol - minVol);

            if (volRange == 0) volRange = 1.0; // 防止除以零

            float yDistance = yAxisLength / (float)volRange;

            // 繪製最頂部與最底部的刻度文字
            using (System.Drawing.Font font = new System.Drawing.Font(this.Font.FontFamily, 6))
            {
                g.DrawString(maxVol.ToString("N0"), font, Brushes.Black, 10, (int)frameTopY);
                g.DrawString(minVol.ToString("N0"), font, Brushes.Black, 10, (int)frameBottomY - 10);

                // 繪製 3 條參考虛線與 Y 軸標籤
                using (Pen dashPen = new Pen(Color.Black, 1) { DashStyle = DashStyle.Dash })
                {
                    double stepVol = volRange / 4.0;
                    for (int i = 1; i <= 3; i++)
                    {
                        float yPos = frameBottomY - (yAxisLength * i / 4f);
                        PointF pl = new PointF(FrameLeftPoints[framePos].frameX, yPos);
                        PointF pr = new PointF(FrameMiddlePoints[framePos].frameX, yPos);

                        g.DrawLine(dashPen, pl, pr);

                        double labelVol = minVol + (stepVol * i);
                        // "N0" : 1234 → 1,234，且不顯示小數
                        g.DrawString(labelVol.ToString("N0"), font, Brushes.Black, 10, (int)yPos - 6);
                    }
                }
            }

            // 準備繪製成交量柱狀圖 (共用 Brush 避免記憶體洩漏)
            using (Brush redBrush = new SolidBrush(Color.Red))
            using (Brush greenBrush = new SolidBrush(Color.Green))
            {
                // 修正 X 座標起始位置，使用對應 framePos 的邊界
                float barXCoord = FrameLeftPoints[framePos].frameX;

                for (int i = StartIndex; i < (StartIndex + DisplayCount); i++)
                {
                    if (i >= StkData.Length) break;

                    if (i != StartIndex)
                    {
                        barXCoord += FrameBarWidth;
                    }

                    double currentVol = StkData[i].Volume;
                    float barHeight = (float)((currentVol - minVol) * yDistance);
                    float barYCoord = frameBottomY - barHeight;

                    // 開盤 > 收盤 為跌(綠)，否則為漲/平(紅)
                    Brush currentBrush = (StkData[i].StartPrice > StkData[i].EndPrice) ? greenBrush : redBrush;

                    g.FillRectangle(currentBrush, barXCoord, barYCoord, FrameBarWidth - 1, barHeight);
                }
            }
            Debug.WriteLine("Chalk_MAP_VOLUME() END!!!!!");
        }

        // 繪製以中心線為基準的單線圖 (ex: MAP_BIAS)
        private void Chalk_MAP_CENTER_SINGLE_LINE(Graphics g, int framePos, int frameNum, int mapType)
        {
            if (framePos <= 0 || framePos > frameNum)
                return;

            float xAxisLength = FrameMiddlePoints[framePos].frameX - FrameLeftPoints[framePos].frameX;      // 儲存要繪製指標柱狀圖的X軸長度
            float yAxisHeight = FrameLeftPoints[framePos].frameY - FrameLeftPoints[framePos - 1].frameY;    // 儲存要繪製指標柱狀圖的Y軸長度
            float yDistance = yAxisHeight / 4;

            HighLowValues highLowValues = GeneralModule.GetHighLowValue(StkData, IdxData, StartIndex, DisplayCount, GeneralModule.MAP_BIAS);
            Debug.WriteLine("最高/低價(BIAS)：" + $"{highLowValues.highValue}, {highLowValues.lowValue}");

            // 劃虛線 (劃3條)
            using (Pen pen = new Pen(Color.Black, 1))
            {
                pen.DashStyle = DashStyle.Dash;
                int p = -1;
                float ff = (float)highLowValues.highValue / 2f;
                for (int i = 1; i <= 3; i++)
                {
                    PointF pl = new PointF(FrameLeftPoints[framePos].frameX, FrameLeftPoints[framePos].frameY - yDistance * i);
                    PointF pr = new PointF(FrameMiddlePoints[framePos].frameX, FrameMiddlePoints[framePos].frameY - yDistance * i);
                    g.DrawLine(pen, pl, pr);
                    Debug.WriteLine("LABEL (BIAS)：" + $"{pl}, {pr}");
                    g.DrawString($"{(ff * (p++))}", new System.Drawing.Font(this.Font.FontFamily, 6), Brushes.Black, 10, (int)(pl.Y));
                }
            }

            // Draw Index Values
            PointF[] points = new PointF[DisplayCount];
            float xWidth = xAxisLength / DisplayCount;
            float yHeight = yAxisHeight / (Math.Abs((float)highLowValues.highValue) * 2f);
            Debug.WriteLine("xWidth= " + xWidth + ", yHeight= " + yHeight);

            float xCoord = FrameLeftPoints[framePos].frameX + (xWidth / 2f);
            float yCoord = FrameLeftPoints[framePos].frameY;
            //Debug.WriteLine("Baseline -- framePos= " + framePos +
            //    ", frame[" + framePos + "].X= " + FrameLeftPoints[framePos].frameX +
            //    ", frame[" + framePos + "].Y= " + FrameLeftPoints[framePos].frameY +
            //    ", xCoord= " + xCoord + ", yCoord= " + yCoord);

            for (int i = StartIndex; i < (StartIndex + DisplayCount); i++)
            {
                float ii = 0;
                if (i == StartIndex)
                {
                    if (IdxData[StartIndex].BIAS >= 0)
                    {
                        ii = ((float)IdxData[StartIndex].BIAS + Math.Abs((float)highLowValues.highValue)) * yHeight;
                    }
                    else
                    {
                        ii = Math.Abs((float)IdxData[StartIndex].BIAS) * yHeight;
                    }
                    yCoord = (float)FrameLeftPoints[framePos].frameY - ii;
                    points[0] = new PointF(xCoord, yCoord);
                }
                else
                {
                    xCoord += xWidth;
                    if (IdxData[i].BIAS >= 0)
                    {
                        ii = ((float)IdxData[i].BIAS + Math.Abs((float)highLowValues.highValue)) * yHeight;
                    }
                    else
                    {
                        ii = Math.Abs((float)IdxData[i].BIAS) * yHeight;
                    }
                    yCoord = (float)FrameLeftPoints[framePos].frameY - ii;
                    points[i - StartIndex] = new PointF(xCoord, yCoord);
                }
                Debug.WriteLine("BIAS[" + i + "] -- xCoord= " + xCoord + ", yCoord= " + yCoord +
                    ", IdxData[i].BIAS= " + IdxData[i].BIAS + ", frameY= " + (float)FrameLeftPoints[framePos].frameY);
            }

            using (Pen pen = new Pen(Color.Blue, 1))
            {
                g.DrawLines(pen, points);
            }
            Debug.WriteLine("Chalk_MAP_BIAS() END!!!!!");
        }

        // 已廢棄，改用 Chalk_MAP_SINGLE_LINE_2()
        private void Chalk_MAP_SINGLE_LINE(Graphics g, int framePos, int frameNum, int mapType)
        {
            if (framePos <= 0 || framePos > frameNum || DisplayCount <= 0)
                return;

            float xAxisLength = FrameMiddlePoints[framePos].frameX - FrameLeftPoints[framePos].frameX;
            float yAxisHeight = FrameLeftPoints[framePos].frameY - FrameLeftPoints[framePos - 1].frameY;
            float yDistance = yAxisHeight / 4f;

            // 繪製 0~100 指標常見的 25%, 50%, 75% 參考虛線與標籤
            using (Pen pen = new Pen(Color.Black, 1) { DashStyle = DashStyle.Dash, DashPattern = new float[] { 7, 3 } })
            using (System.Drawing.Font font = new System.Drawing.Font(this.Font.FontFamily, 6))
            {
                for (int i = 1; i <= 3; i++)
                {
                    float yPos = FrameLeftPoints[framePos].frameY - yDistance * i;
                    PointF pl = new PointF(FrameLeftPoints[framePos].frameX, yPos);
                    PointF pr = new PointF(FrameMiddlePoints[framePos].frameX, yPos);
                    g.DrawLine(pen, pl, pr);

                    // 固定顯示 25, 50, 75 刻度（因總高為 100）
                    int labelValue = i * 25;
                    g.DrawString(labelValue.ToString(), font, Brushes.Black, 10, yPos - 6);
                }
            }

            // 獲取指標數據
            double[] values = mapType switch
            {
                GeneralModule.MAP_WMS => IdxData.Select(d => d.WMS).ToArray(),
                GeneralModule.MAP_PSY => IdxData.Select(d => d.PSY).ToArray(),
                GeneralModule.MAP_SRSI => IdxData.Select(d => d.SRSI).ToArray(),
                GeneralModule.MAP_MARGIN_PURCHASE => StkData.Select(d => d.MarginPurchase).ToArray(),
                GeneralModule.MAP_SHORT_SELLING => StkData.Select(d => d.ShortSelling).ToArray(),
                _ => new double[IdxData.Count()]
            };

            // 計算各數據點座標
            PointF[] points = new PointF[DisplayCount];
            float xWidth = xAxisLength / DisplayCount;
            float yHeight = yAxisHeight / 100f; // 100 分制轉換比率
            float startX = FrameLeftPoints[framePos].frameX + (xWidth / 2f);
            float baseFrameY = FrameLeftPoints[framePos].frameY;

            for (int i = 0; i < DisplayCount; i++)
            {
                int dataIdx = StartIndex + i;
                if (dataIdx >= values.Length) break;

                float xCoord = startX + (i * xWidth);
                float yCoord = baseFrameY - (float)(values[dataIdx] * yHeight);
                points[i] = new PointF(xCoord, yCoord);
            }

            // 繪製折線
            using (Pen pen = new Pen(Color.Green, 1))
            {
                g.DrawLines(pen, points);
            }
        }

        //  
        private void Chalk_MAP_DOUBLE_LINE(Graphics g, int framePos, int frameNum, int mapType1, int mapType2)
        {
            if (framePos <= 0 || framePos > frameNum || DisplayCount <= 0)
                return;

            float xAxisLength = FrameMiddlePoints[framePos].frameX - FrameLeftPoints[framePos].frameX;
            float yAxisHeight = FrameLeftPoints[framePos].frameY - FrameLeftPoints[framePos - 1].frameY;
            float yDistance = yAxisHeight / 4f;

            // 繪製參考虛線
            using (Pen pen = new Pen(Color.Black, 1) { DashStyle = DashStyle.Dash, DashPattern = new float[] { 7, 3 } })
            using (System.Drawing.Font font = new System.Drawing.Font(this.Font.FontFamily, 6))
            {
                for (int i = 1; i <= 3; i++)
                {
                    float yPos = FrameLeftPoints[framePos].frameY - yDistance * i;
                    PointF pl = new PointF(FrameLeftPoints[framePos].frameX, yPos);
                    PointF pr = new PointF(FrameMiddlePoints[framePos].frameX, yPos);
                    g.DrawLine(pen, pl, pr);
                    g.DrawString((i * 25).ToString(), font, Brushes.Black, 10, yPos - 6);
                }
            }

            // 獲取雙線數據
            double[] values1 = new double[IdxData.Count()];
            double[] values2 = new double[IdxData.Count()];

            if (mapType1 == GeneralModule.MAP_KD || mapType2 == GeneralModule.MAP_KD)
            {
                values1 = IdxData.Select(d => d.K).ToArray();
                values2 = IdxData.Select(d => d.D).ToArray();
            }
            else if (mapType1 == GeneralModule.MAP_RSI || mapType2 == GeneralModule.MAP_RSI)
            {
                values1 = IdxData.Select(d => d.SRSI).ToArray();
                values2 = IdxData.Select(d => d.LRSI).ToArray();
            }
            else
            {
                values1 = GetSingleIndicatorArray(mapType1);
                values2 = GetSingleIndicatorArray(mapType2);
            }

            // 計算兩條折線節點
            PointF[] points1 = new PointF[DisplayCount];
            PointF[] points2 = new PointF[DisplayCount];
            float xWidth = xAxisLength / DisplayCount;
            float yHeight = yAxisHeight / 100f;
            float startX = FrameLeftPoints[framePos].frameX + (xWidth / 2f);
            float baseFrameY = FrameLeftPoints[framePos].frameY;

            for (int i = 0; i < DisplayCount; i++)
            {
                int dataIdx = StartIndex + i;
                if (dataIdx >= IdxData.Count()) break;

                float xCoord = startX + (i * xWidth);
                float y1 = baseFrameY - (float)(values1[dataIdx] * yHeight);
                float y2 = baseFrameY - (float)(values2[dataIdx] * yHeight);

                points1[i] = new PointF(xCoord, y1);
                points2[i] = new PointF(xCoord, y2);
            }

            // 畫線
            using (Pen pen1 = new Pen(Color.Green, 1))
            {
                g.DrawLines(pen1, points1);
            }
            using (Pen pen2 = new Pen(Color.RosyBrown, 1) { DashStyle = DashStyle.Dash, DashPattern = new float[] { 6, 2 } })
            {
                g.DrawLines(pen2, points2);
            }
        }

        // 輔助方法：根據指標類型取得單一數值陣列
        private double[] GetSingleIndicatorArray(int mapType)
        {
            return mapType switch
            {
                GeneralModule.MAP_WMS => IdxData.Select(d => d.WMS).ToArray(),
                GeneralModule.MAP_PSY => IdxData.Select(d => d.PSY).ToArray(),
                GeneralModule.MAP_SRSI => IdxData.Select(d => d.SRSI).ToArray(),
                GeneralModule.MAP_LRSI => IdxData.Select(d => d.LRSI).ToArray(),
                GeneralModule.MAP_K => IdxData.Select(d => d.K).ToArray(),
                GeneralModule.MAP_D => IdxData.Select(d => d.D).ToArray(),
                _ => new double[IdxData.Count()]
            };
        }

        //  
        private void Chalk_MAP_MACD_LINE(Graphics g, int framePos, int frameNum, int mapType)
        {
            if (framePos <= 0 || framePos > frameNum || DisplayCount <= 0)
                return;

            float xAxisLength = FrameMiddlePoints[framePos].frameX - FrameLeftPoints[framePos].frameX;
            float yAxisHeight = FrameLeftPoints[framePos].frameY - FrameLeftPoints[framePos - 1].frameY;
            float yDistance = yAxisHeight / 4f;

            HighLowValues highLowValues = GeneralModule.GetHighLowValue(StkData, IdxData, StartIndex, DisplayCount, mapType);

            // 計算對稱零軸的最大絕對值
            double maxV = Math.Max(Math.Abs(highLowValues.highValue), Math.Abs(highLowValues.lowValue));
            if (maxV == 0) maxV = 1.0; // 防止除以零

            double[] LV = new double[3] { maxV / 2.0, 0, -maxV / 2.0 }; // 從上到下: 正、零、負

            // 劃 3 條參考虛線與 Y 軸文字
            using (System.Drawing.Font font = new System.Drawing.Font(this.Font.FontFamily, 6))
            using (Pen dashPen = new Pen(Color.Black, 1) { DashStyle = DashStyle.Dash, DashPattern = new float[] { 7, 3 } })
            {
                for (int i = 1; i <= 3; i++)
                {
                    float yPos = FrameLeftPoints[framePos].frameY - yDistance * i;
                    PointF pl = new PointF(FrameLeftPoints[framePos].frameX, yPos);
                    PointF pr = new PointF(FrameMiddlePoints[framePos].frameX, yPos);

                    g.DrawLine(dashPen, pl, pr);
                    g.DrawString(LV[i - 1].ToString("0.00"), font, Brushes.Black, 10, yPos - 6);
                }
            }

            // 比例轉換與數據準備
            float xWidth = xAxisLength / DisplayCount;
            float yHeight = yAxisHeight / (float)(maxV * 2.0); // 數值到像素的轉換比率
            float zeroY = FrameLeftPoints[framePos].frameY - (yAxisHeight / 2.0f); // 零軸 Y 座標 (畫布中央)

            double[] difValues = IdxData.Select(d => d.DIF).ToArray();
            double[] macdValues = IdxData.Select(d => d.MACD).ToArray();
            double[] oscValues = IdxData.Select(d => d.OSC).ToArray();

            PointF[] difPoints = new PointF[DisplayCount];
            PointF[] macdPoints = new PointF[DisplayCount];

            // 計算折線點座標
            for (int i = 0; i < DisplayCount; i++)
            {
                int dataIdx = StartIndex + i;
                if (dataIdx >= IdxData.Count()) break;

                float currentX = FrameLeftPoints[framePos].frameX + (i * xWidth) + (xWidth / 2f);

                float difY = zeroY - (float)(difValues[dataIdx] * yHeight);
                float macdY = zeroY - (float)(macdValues[dataIdx] * yHeight);

                difPoints[i] = new PointF(currentX, difY);
                macdPoints[i] = new PointF(currentX, macdY);
            }

            // 繪製 MACD 柱狀圖 (Oscillator / Histogram)
            for (int i = 0; i < DisplayCount; i++)
            {
                int dataIdx = StartIndex + i;
                if (dataIdx >= IdxData.Count()) break;

                double val = oscValues[dataIdx];
                float barHeight = (float)(Math.Abs(val) * yHeight);
                float xPos = FrameLeftPoints[framePos].frameX + (i * xWidth);

                if (val >= 0)
                {
                    // 正值：由零軸往上畫
                    using (Brush redBrush = new SolidBrush(Color.Red))
                    {
                        g.FillRectangle(redBrush, xPos, zeroY - barHeight, xWidth - 1, barHeight);
                    }
                }
                else
                {
                    // 負值：由零軸往下畫
                    using (Brush greenBrush = new SolidBrush(Color.Green))
                    {
                        g.FillRectangle(greenBrush, xPos, zeroY, xWidth - 1, barHeight);
                    }
                }
            }

            // 繪製 DIF 與 DEA 兩條線
            using (Pen penDIF = new Pen(Color.Green, 1))
            {
                g.DrawLines(penDIF, difPoints);
            }

            using (Pen penMACD = new Pen(Color.RosyBrown, 1) { DashStyle = DashStyle.Dash, DashPattern = new float[] { 6, 2 } })
            {
                g.DrawLines(penMACD, macdPoints);
            }
        }

        // 輔助方法：繪製 MACD 柱狀圖 (Oscillator / Histogram) 與 DIF/DEA 線
        private void Chalk_MAP_CENTER_BAR(Graphics g, int framePos, int frameNum, int mapType)
        {
            // 邊界與參數安全檢查
            if (framePos <= 0 || framePos > frameNum || DisplayCount <= 0)
                return;

            float xAxisLength = FrameMiddlePoints[framePos].frameX - FrameLeftPoints[framePos].frameX;
            float yAxisHeight = FrameLeftPoints[framePos].frameY - FrameLeftPoints[framePos - 1].frameY;
            float yDistance = yAxisHeight / 4f;

            HighLowValues highLowValues = GeneralModule.GetHighLowValue(StkData, IdxData, StartIndex, DisplayCount, mapType);

            // 計算對稱零軸的最大絕對值
            double maxV = Math.Max(Math.Abs(highLowValues.highValue), Math.Abs(highLowValues.lowValue));
            if (maxV == 0) maxV = 1.0; // 防止除以零

            double[] LV = new double[3] { maxV / 2.0, 0, -maxV / 2.0 }; // 從上到下: 正、零、負

            // 劃 3 條參考虛線與 Y 軸文字
            using (System.Drawing.Font font = new System.Drawing.Font(this.Font.FontFamily, 6))
            using (Pen dashPen = new Pen(Color.Black, 1) { DashStyle = DashStyle.Dash, DashPattern = new float[] { 7, 3 } })
            {
                for (int i = 1; i <= 3; i++)
                {
                    float yPos = FrameLeftPoints[framePos].frameY - yDistance * i;
                    PointF pl = new PointF(FrameLeftPoints[framePos].frameX, yPos);
                    PointF pr = new PointF(FrameMiddlePoints[framePos].frameX, yPos);

                    g.DrawLine(dashPen, pl, pr);
                    g.DrawString(LV[i - 1].ToString("0.00"), font, Brushes.Black, 10, yPos - 6);
                }
            }

            // 比例轉換與數據準備
            float xWidth = xAxisLength / DisplayCount;
            float yHeight = yAxisHeight / (float)(maxV * 2.0); // 數值到像素的轉換比率
            float zeroY = FrameLeftPoints[framePos].frameY - (yAxisHeight / 2.0f); // 零軸 Y 座標 (畫布中央)
            double[] values = IdxData.Select(d => d.SECTORS).ToArray();

            // 繪製 MACD 柱狀圖 (Histogram)
            for (int i = 0; i < DisplayCount; i++)
            {
                int dataIdx = StartIndex + i;
                if (dataIdx >= IdxData.Count()) break;

                double val = values[dataIdx];
                float barHeight = (float)(Math.Abs(val) * yHeight);
                float xPos = FrameLeftPoints[framePos].frameX + (i * xWidth);

                if (val >= 0)
                {
                    // 正值：由零軸往上畫
                    using (Brush redBrush = new SolidBrush(Color.Red))
                    {
                        g.FillRectangle(redBrush, xPos, zeroY - barHeight, xWidth - 1, barHeight);
                    }
                }
                else
                {
                    // 負值：由零軸往下畫
                    using (Brush greenBrush = new SolidBrush(Color.Green))
                    {
                        g.FillRectangle(greenBrush, xPos, zeroY, xWidth - 1, barHeight);
                    }
                }
            }
            Debug.WriteLine("Chalk_MAP_VOLUME() END!!!!!");
        }

        // 助手方法：繪製單線圖 (MAP_WMS, MAP_PSY, MAP_SRSI, MAP_MARGIN_PURCHASE, MAP_SHORT_SELLING)
        private void Chalk_MAP_SINGLE_LINE_2(Graphics g, int framePos, int frameNum, int mapType)
        {
            if (framePos <= 0 || framePos > frameNum || DisplayCount <= 0)
                return;

            float frameLeftX = FrameLeftPoints[framePos].frameX;
            float frameRightX = FrameMiddlePoints[framePos].frameX;
            float frameTopY = FrameLeftPoints[framePos - 1].frameY;
            float frameBottomY = FrameLeftPoints[framePos].frameY;

            float xAxisLength = frameRightX - frameLeftX;   // X 軸長度
            float yAxisHeight = frameBottomY - frameTopY;   // Y 軸高度

            HighLowValues range = GeneralModule.GetHighLowValue(StkData, IdxData, StartIndex, DisplayCount, mapType);

            double maxValue = range.highValue;
            double minValue = range.lowValue;

            double classWidth = maxValue - minValue;

            if (classWidth <= 0)
                classWidth = 1;

            float valueToPixel = yAxisHeight / (float)classWidth;   // 每個像素的高度

            // 畫3條虛線及最左側的值(5個)
            DrawReferenceLines(g, frameLeftX, frameRightX, frameBottomY, yAxisHeight, minValue, maxValue);

            double[] values = mapType switch
            {
                GeneralModule.MAP_WMS => IdxData.Select(d => d.WMS).ToArray(),
                GeneralModule.MAP_PSY => IdxData.Select(d => d.PSY).ToArray(),
                GeneralModule.MAP_SRSI => IdxData.Select(d => d.SRSI).ToArray(),
                GeneralModule.MAP_MARGIN_PURCHASE => StkData.Select(d => d.MarginPurchase).ToArray(),
                GeneralModule.MAP_SHORT_SELLING => StkData.Select(d => d.ShortSelling).ToArray(),
                _ => Array.Empty<double>()
            };

            // 劃指標線
            DrawIndicatorLine(g, values, frameLeftX, frameBottomY, xAxisLength, valueToPixel, minValue);
        }

        // 助手方法：繪製參考虛線與標籤
        private void DrawReferenceLines(Graphics g, float leftX, float rightX, float bottomY, float height, double minValue, double maxValue)
        {
            using Pen pen = new(Color.Black, 1)
            {
                DashStyle = DashStyle.Dash,
                DashPattern = new float[] { 7, 3 }
            };

            using System.Drawing.Font font = new(this.Font.FontFamily, 6);

            double range = maxValue - minValue;

            for (int i = 0; i <= 4; i++)
            {
                float y = bottomY - height * i / 4f;
                double value = minValue + range * i / 4.0;
                if (i == 4)
                    g.DrawString(((int)value).ToString(), font, Brushes.Black, 10, y + 6);
                else
                    g.DrawString(((int)value).ToString(), font, Brushes.Black, 10, y - 6);

                if (i > 0 && i < 4)
                {
                    g.DrawLine(pen, leftX, y, rightX, y);
                }
            }
        }

        // 助手方法：繪製指標折線
        private void DrawIndicatorLine(Graphics g, double[] values, float leftX, float bottomY, float width, float valueToPixel, double minValue)
        {
            List<PointF> points = new();
            float xStep = width / DisplayCount;
            float startX = leftX + xStep / 2f;

            for (int i = 0; i < DisplayCount; i++)
            {
                int dataIdx = StartIndex + i;

                if (dataIdx >= values.Length)
                    break;

                float x = startX + i * xStep;
                float y = bottomY - (float)((values[dataIdx] - minValue) * valueToPixel);

                points.Add(new PointF(x, y));
            }

            if (points.Count < 2)
                return;

            using Pen pen = new(Color.Green, 1);

            g.DrawLines(pen, points.ToArray());
        }











        //-- Write Next Here --//



    }
}
