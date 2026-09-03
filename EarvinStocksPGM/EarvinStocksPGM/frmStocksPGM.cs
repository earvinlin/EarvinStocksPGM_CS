using MySqlConnector;
using System.Diagnostics;
using System.Drawing.Drawing2D;

namespace EarvinStocksPGM
{
    struct FramePoints
    {
        public float frameX;
        public float frameY;
    }

    public partial class frmStocksPGM : Form
    {
        private int frameNum = 0;               // frame數量
        private float XWidthBorder = 20;        // frame左、右兩邊預留的空間
        private float YHeightBorder = 10;       // frame最下面預留的空間
        private float frmXTop = 30;             // frame最左上角的X座標
        private float frmRightBorder = 150;     // frame最左上角的Y座標

        private Label lblStokInfo;              // 動態新增label元件：顯示股票資訊用
        private Label lblHighPrice;             // 動態新增label元件：顯示股票最高價
        private Label lblLowPrice;              // 動態新增label元件：顯示股票最低價
        private Label[] lblStockYM = new Label[12];

        private Boolean blnShowFocusLine = false;  // 是否顯示焦點線段
        int displayCount = 100; // 顯示的資料筆數
        int startIndex = 0; // 顯示的資料起始索引

        private Point cursorPosition = new Point(); // 滑鼠游標位置
        private Point lineFocusTop = new Point();
        private Point lineFocusBottom = new Point();
        float frmTopXCoord = 0;
        float frmTopYCoord = 0;
        float frmBottomXCoord = 0;
        float frmBottomYCoord = 0;
        float frmXAxisWidth = 0;
        float frmBarWidth = 0;



        public frmStocksPGM()
        {
            InitializeComponent();
            this.DoubleBuffered = true;

            this.StartPosition = FormStartPosition.CenterScreen;
            pnlStocksBar.Width = this.Width;
        }
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            // TODO
        }

        private void pnlStocksBar_Paint(object sender, PaintEventArgs e)
        {
        }

        private void cboStocks_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void label2_Click(object sender, EventArgs e)
        {
        }

        private void btnFocus_Click(object sender, EventArgs e)
        {
            blnShowFocusLine = !blnShowFocusLine;
        }

        private void frmStocksPGM_Load(object sender, EventArgs e)
        {
            pnlStocksBar.Width = this.Width;
            frameNum = int.Parse(cboFrameNum.Text);

            // 新增顯示股票資訊的標籤
            lblStokInfo = new Label()
            {
                Name = "lblStokInfo",
                Text = "This is a test message!",
                AutoSize = true,
                Location = new Point(10, mnuStocksList.Size.Height + pnlStocksBar.Size.Height)
            };
            this.Controls.Add(lblStokInfo);

            // 新增顯示股票資訊的標籤
            lblHighPrice = new Label()
            {
                Name = "lblHighPrice",
                Text = "high",
                AutoSize = true,
                Font = new Font(this.Font.FontFamily, 8),
                //Location = new Point(10, mnuStocksList.Size.Height + pnlStocksBar.Size.Height)
            };
            this.Controls.Add(lblHighPrice);

            // 新增顯示股票資訊的標籤
            lblLowPrice = new Label()
            {
                Name = "lblLowPrice",
                Text = "low",
                AutoSize = true,
                Font = new Font(this.Font.FontFamily, 8),
                //Location = new Point(10, mnuStocksList.Size.Height + pnlStocksBar.Size.Height)
            };
            this.Controls.Add(lblLowPrice);

            // 新增 Label 元件(預設建立12個備用)
            //Label[] lblStockYM = new Label[12];
            for (int i = 0; i < 12; i++)
            {
                lblStockYM[i] = new Label()
                {
                    Name = $"lblStockYM{i}",
                    Text = "YYMM",
                    AutoSize = true,
                    Font = new Font(this.Font.FontFamily, 8),
                };
                this.Controls.Add(lblStockYM[i]);
                Debug.WriteLine("lblStockYM[" + i + "]：" + lblStockYM[i].Text);
            }
        }

        private void frmStocksPGM_Resize(object sender, EventArgs e)
        {
            pnlStocksBar.Width = this.Width;
            this.Invalidate();
        }

        private void frmStocksPGM_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            e.Graphics.Clear(this.BackColor);

            // lblStokInfo : 顯示股票資訊
            float frmYTop = mnuStocksList.Size.Height + pnlStocksBar.Size.Height + lblStokInfo.Size.Height;
            // XWidthBorder : 表示frame左右皆各內縮 (XWidthBorder / 2) 個pixels
            float frmXWidth = this.ClientSize.Width - frmXTop - (XWidthBorder / 2);
            // YHeightBorder : 表示frame最下面上調YHeightBorder個pixels
            float frmYHeight = this.ClientSize.Height - frmYTop - YHeightBorder;
            // Frame 數量
            frameNum = int.Parse(cboFrameNum.Text);
            //Debug.WriteLine("frameNum= " + frameNum);

            //---------------------//
            //-- 繪製 Frame 外框 --//
            //---------------------//
            FramePoints[] frameLeftPoints = new FramePoints[frameNum + 1];  // FramePoints結構陣列，存放frame左邊各個點的座標
            FramePoints[] frameRightPoints = new FramePoints[frameNum + 1]; // FramePoints結構陣列，存放frame右邊各個點的座標

            for (int i = 0; i < (frameNum + 1); i++)
            {
                frameLeftPoints[i].frameX = frmXTop;
                if (i == 0)
                {
                    frameLeftPoints[i].frameY = frmYTop;
                }
                else if (i == 1)
                {
                    frameLeftPoints[i].frameY = frmYTop + frmYHeight / 2;
                }
                else
                {
                    frameLeftPoints[i].frameY = frmYTop + (frmYHeight / 2) + (frmYHeight / 2) / (frameNum - 1) * (i - 1);
                }
                // (For CHECK / DEBUG) 顯示Frame最左側端點座標
                g.FillEllipse(Brushes.BlueViolet, frameLeftPoints[i].frameX, frameLeftPoints[i].frameY, 5, 5);
            }

            for (int i = 0; i < (frameNum + 1); i++)
            {
                frameRightPoints[i].frameX = frmXTop + frmXWidth;
                if (i == 0)
                {
                    frameRightPoints[i].frameY = frmYTop;
                }
                else if (i == 1)
                {
                    frameRightPoints[i].frameY = frmYTop + frmYHeight / 2;
                }
                else
                {
                    frameRightPoints[i].frameY = frmYTop + (frmYHeight / 2) + (frmYHeight / 2) / (frameNum - 1) * (i - 1);
                }
                // (For CHECK / DEBUG) 顯示Frame最右側端點座標
                g.FillEllipse(Brushes.BlueViolet, frameRightPoints[i].frameX, frameRightPoints[i].frameY, 5, 5);
            }

            // FramePoints結構陣列，存放frame中間各個點的座標(最後1個點是frame最右下角的點)
            FramePoints[] frameMiddlePoints = new FramePoints[frameNum + 1];
            for (int i = 0; i < (frameNum + 1); i++)
            {
                frameMiddlePoints[i].frameX = frmXTop + (frmXWidth - frmRightBorder);
                if (i == 0)
                {
                    frameMiddlePoints[i].frameY = frmYTop;
                }
                else if (i == 1)
                {
                    frameMiddlePoints[i].frameY = frmYTop + frmYHeight / 2;
                }
                else if (i == frameNum)
                {
                    frameMiddlePoints[i].frameY = frmYTop + frmYHeight;
                }
                else
                {
                    frameMiddlePoints[i].frameY = frmYTop + (frmYHeight / 2) + (frmYHeight / 2) / (frameNum - 1) * (i - 1);
                }
                // (For CHECK / DEBUG) 顯示Frame內側端點座標
                g.FillEllipse(Brushes.BlueViolet, frameMiddlePoints[i].frameX, frameMiddlePoints[i].frameY, 5, 5);
            }
            // Frame外框
            g.DrawRectangle(Pens.Blue, frmXTop, frmYTop, frmXWidth, frmYHeight);
            // 每個Frame的分隔線
            g.DrawLine(Pens.Magenta, frameMiddlePoints[0].frameX, frameMiddlePoints[0].frameY, frameMiddlePoints[frameNum].frameX, frameMiddlePoints[frameNum].frameY);
            for (int i = 1; i < frameNum; i++)
            {
                g.DrawLine(Pens.Brown, frameLeftPoints[i].frameX, frameLeftPoints[i].frameY, frameRightPoints[i].frameX, frameRightPoints[i].frameY);
            }

            //--------------------------------------------//
            //-- 繪製 K-Map (最上方 Frame) 的橫(虛)線段 --//
            //--------------------------------------------//
            int dashLineCounts = 5; // 虛線的段數
            float hh = frameLeftPoints[1].frameY - frameLeftPoints[0].frameY;

            Pen pen = new Pen(Color.Black, 1);
            pen.DashStyle = DashStyle.Dash;
            pen.DashPattern = new float[] { 4, 2 }; // 畫 4px，空 2px
            for (int i = 1; i < dashLineCounts; i++)
            {
                float x0 = frameLeftPoints[0].frameX;
                float y0 = frameLeftPoints[0].frameY + (hh / dashLineCounts) * i;
                float x1 = frameMiddlePoints[0].frameX;
                float y1 = frameMiddlePoints[0].frameY + (hh / dashLineCounts) * i;
                g.DrawLine(pen, x0, y0, x1, y1);
            }

            // 取得要顯示的股票資料
            DbHelper.StockData[] sd = DbHelper.TestConnectDB();
            if (sd == null || sd.Length == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }
            if (startIndex > (sd.Length - displayCount))
            {
                startIndex = sd.Length - displayCount;
            }
            if (startIndex < 0)
            {
                startIndex = 0;
            }

            // 顯示畫面筆數之最高/最低價 (因為資料庫的資料型態為 decimal，為了便於計算故宣告為 decimal)
            decimal stockPriceHighest = 0;
            decimal stockProceLowest = 99999;

            Debug.WriteLine($"sd counts = {sd.Length}, startIndex = {startIndex}");
            for (int i = startIndex; i < (startIndex + displayCount - 1); i++)
            {
                if (stockPriceHighest < sd[i].HighPrice)
                    stockPriceHighest = sd[i].HighPrice;
                if (stockProceLowest > sd[i].LowPrice)
                    stockProceLowest = sd[i].LowPrice;
            }
            Debug.WriteLine("最高/低價：" + $"{stockPriceHighest}, {stockProceLowest}");

            //--------------------------------------//
            //-- 顯示 K-Map (最上方 Frame) 柱狀圖 --//
            //--------------------------------------//
            float XAxisLength = frameMiddlePoints[0].frameX - frameLeftPoints[0].frameX;            // 顯示 K-Map's Frame 的X軸長度
            float YAxisLength = frameLeftPoints[1].frameY - frameLeftPoints[0].frameY;              // 顯示 K-Map's Frame 的Y軸長度
            float barWidth = XAxisLength / (float)displayCount;    // 要繪製K-Bar的寬度
            float barHeight = 0;                            // 要繪製K-Bar的高度
            float barXCoord = frameLeftPoints[0].frameX;    // 要繪製K-Bar的X座標
            float barYCoord = 0;                            // 要繪製K-Bar的Y座標
            float yDistance = YAxisLength / (float)Math.Abs(stockPriceHighest - stockProceLowest); // 取得每個價格對應的Y軸距離
            //Debug.WriteLine("yDistance= " + yDistance);

//            for (int i = startIndex; i < (startIndex + displayCount - 1); i++)
              for (int i = startIndex; i < (startIndex + displayCount); i++)
                {
                    // X 座標
//                    if (i != 0)
                    if (i != startIndex)
                    {
                    barXCoord += barWidth;
                }
                // Y 座標
                if (sd[i].StartPrice > sd[i].EndPrice)
                    barYCoord = (float)frameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(stockPriceHighest - sd[i].StartPrice));
                else
                    barYCoord = (float)frameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(stockPriceHighest - sd[i].EndPrice));
                // 計算 K-Bar 的高度
                barHeight = yDistance * (float)Math.Abs(sd[i].StartPrice - sd[i].EndPrice);
                // 繪製 K-Bar
                if (barHeight != 0)
                {
                    if (sd[i].StartPrice > sd[i].EndPrice)
                    {
                        Brush brush = new SolidBrush(Color.Green);
                        g.FillRectangle(brush, barXCoord, barYCoord, barWidth, barHeight);
                    }
                    else
                    {
                        Brush brush = new SolidBrush(Color.Red);
                        g.FillRectangle(brush, barXCoord, barYCoord, barWidth, barHeight);
                    }
                }
                else
                {
                    g.DrawLine(Pens.Black, barXCoord, barYCoord, (barXCoord + barWidth), barYCoord);
                }
                // 繪製 K-Bar 最高價 to 最低價之線段
                float x0 = (barXCoord + barWidth / 2);
                float y0 = (float)frameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(stockPriceHighest - sd[i].HighPrice));
                float x1 = (barXCoord + barWidth / 2);
                float y1 = (float)frameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(stockPriceHighest - sd[i].LowPrice));

                if (sd[i].StartPrice > sd[i].EndPrice)
                    g.DrawLine(Pens.Green, x0, y0, x1, y1);
                else
                    g.DrawLine(Pens.Red, x0, y0, x1, y1);
                //              Debug.WriteLine("date= " + sd[i].TradeDate + ", k-bar x0: " + x0 + " ,y0: " + y0, " ,x1-: " + x1 + " ,y1-: " + y1);
            }

            //-------------------------------------//
            //-- 顯示 K-Map (上方Frame) 的直線段 --//
            //-------------------------------------//
            int k = 0;
            long prevNum = 0, nextNum = 0;
//            for (int i = startIndex; i < (startIndex + displayCount - 1); i++)
              for (int i = startIndex; i < (startIndex + displayCount); i++)
                {
                    if (i == 0)
                {
                    // 取交易日期的最後兩碼，若為 20240101，則取 01
                    prevNum = sd[i].TradeDate % 100;
                    nextNum = sd[i].TradeDate % 100;
                    continue;
                }
                nextNum = sd[i].TradeDate % 100;

                if (prevNum > nextNum)
                {
                    float x0 = frameLeftPoints[0].frameX + (barWidth * i);
                    float y0 = frameLeftPoints[0].frameY;
                    float x1 = x0;
                    float y1 = frameLeftPoints[1].frameY;
                    g.DrawLine(pen, x0, y0, x1, y1);
                    Debug.WriteLine("date= " + sd[i].TradeDate + ", k-bar x0: " + x0 + " ,y0: " + y0, " ,x1-: " + x1 + " ,y1-: " + y1);
                    // 顯示交易日期(年月)
                    string strnum = sd[i].TradeDate.ToString();
                    lblStockYM[k].Text = strnum.Substring(0, strnum.Length - 2);
                    lblStockYM[k].Location = new System.Drawing.Point((int)(x0 - (lblStockYM[k].Size.Width / 2)), (int)(y1 + 5));
                    k = k + 1;
                }
                prevNum = nextNum;
            }

            //-- (MouseMvoe Event) ----------------------------------------------------------------------------------------
            int curIndex = 0;
            if (blnShowFocusLine)
            {
                frmTopXCoord = frameLeftPoints[0].frameX;
                frmTopYCoord = frameLeftPoints[0].frameY;
                frmBottomXCoord = frameLeftPoints[frameNum].frameX;
                frmBottomYCoord = frameLeftPoints[frameNum].frameY;
                frmXAxisWidth = XAxisLength;
                frmBarWidth = barWidth;
                g.DrawLine(Pens.Brown, cursorPosition.X, frmTopYCoord, cursorPosition.X, frmBottomYCoord);
                if (cursorPosition.X <= 0)
                    cursorPosition.X = (int)frmTopXCoord;
                if (cursorPosition.X >= (frmTopXCoord + frmXAxisWidth))
                    cursorPosition.X = (int) (frmTopXCoord + frmXAxisWidth);

                else if (cursorPosition.X > (frmTopXCoord + frmXAxisWidth))
                    cursorPosition.X = (int)(frmTopXCoord + frmXAxisWidth);
                curIndex = (int)((cursorPosition.X - frmTopXCoord) / frmBarWidth) + startIndex;
                Debug.WriteLine("AAAA -- curIndex= " + curIndex + ", cursorPosition= " + cursorPosition.X + ", frmTopXCoord= " + frmTopXCoord + ", frmBarWidth= " + frmBarWidth);

            }
            //-------------------------------------------------------------------------------------------------------------

            // Y-Length : 顯示frame的Y軸長度

            lblStokInfo.Text = "日期：" + sd[curIndex].TradeDate + " 開 " + sd[curIndex].StartPrice + " 高 " + sd[curIndex].HighPrice + " 低 " + sd[curIndex].LowPrice + " 收 " + sd[curIndex].EndPrice ;
            lblHighPrice.Location = new System.Drawing.Point((int)frameLeftPoints[0].frameX - lblHighPrice.Width, (int)frameLeftPoints[0].frameY);
            lblLowPrice.Location = new System.Drawing.Point((int)frameLeftPoints[1].frameX - lblLowPrice.Width, (int)frameLeftPoints[1].frameY - lblLowPrice.Height);
        }

        private void cboFrameNum_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Invalidate();
        }

        private void frmStocksPGM_MouseMove(object sender, MouseEventArgs e)
        {
            // 滑鼠移動時，顯示焦點線段
            if (blnShowFocusLine)
            {
                //frmTopXCoord = frameLeftPoints[0].frameX;
                //frmTopYCoord = frameLeftPoints[0].frameY;
                //frmBottomXCoord = frameLeftPoints[frameNum].frameX;
                //frmBottomYCoord = frameLeftPoints[frameNum].frameY;
                //frmXAxisWidth = XAxisLength;
                //frmBarWidth = barWidth;


                //                cursorPosition = Cursor.Position;

                cursorPosition = this.PointToClient(Cursor.Position);
                Debug.WriteLine($"X = {cursorPosition.X}, Y = {cursorPosition.Y}");
                Debug.WriteLine($"frmTopXCoord = {frmTopXCoord}, frmTopYCoord = {frmTopYCoord}");
                Debug.WriteLine($"frmBottomXCoord = {frmBottomXCoord}, frmBottomYCoord = {frmBottomYCoord}");

                // 判斷 FocusLine

                using (Graphics g = this.CreateGraphics())
                {
                    //if (cursorPosition.X < frmTopXCoord)
                    //    g.DrawLine(Pens.Blue, frmTopXCoord, frmTopYCoord, frmTopXCoord, frmBottomYCoord);
                    //else if (cursorPosition.X > (frmTopXCoord + frmXAxisWidth))
                    //    g.DrawLine(Pens.Blue, (frmTopXCoord + frmXAxisWidth), frmTopYCoord, (frmTopXCoord + frmXAxisWidth), frmBottomYCoord);
                    //else
                    //    g.DrawLine(Pens.Blue, cursorPosition.X, frmTopYCoord, cursorPosition.X, frmBottomYCoord);
                    if (cursorPosition.X <= frmTopXCoord)
                        cursorPosition.X = (int)frmTopXCoord;
                    else if (cursorPosition.X > (frmTopXCoord + frmXAxisWidth))
                        cursorPosition.X = (int)(frmTopXCoord + frmXAxisWidth);

                    //g.DrawLine(Pens.Brown, cursorPosition.X, frmTopYCoord, cursorPosition.X, frmBottomYCoord);
                }

                // 觸發重繪
                this.Invalidate();
            }
        }

        private void btnBack1_Click(object sender, EventArgs e)
        {
            startIndex += 1;
            Debug.WriteLine($"startIndex = {startIndex}");
            // 觸發重繪
            this.Invalidate();
        }

        private void btnFore1_Click(object sender, EventArgs e)
        {
            startIndex -= 1;
            Debug.WriteLine($"startIndex = {startIndex}");
            // 觸發重繪
            this.Invalidate();
        }

        private void btnBack2_Click(object sender, EventArgs e)
        {
            startIndex += 5;
            Debug.WriteLine($"startIndex = {startIndex}");
            // 觸發重繪
            this.Invalidate();
        }

        private void btnFore2_Click(object sender, EventArgs e)
        {
            startIndex -= 5;
            Debug.WriteLine($"startIndex = {startIndex}");
            // 觸發重繪
            this.Invalidate();
        }

        private void btnBack3_Click(object sender, EventArgs e)
        {
            startIndex = int.MaxValue;
            Debug.WriteLine($"startIndex = {startIndex}");
            // 觸發重繪
            this.Invalidate();
        }

        private void btnFore3_Click(object sender, EventArgs e)
        {
            startIndex = 0;
            Debug.WriteLine($"startIndex = {startIndex}");
            // 觸發重繪
            this.Invalidate();
        }
    }
}
