using EarvinStocksPGM.Modules;
using MySqlConnector;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using static EarvinStocksPGM.Modules.GeneralModule;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace EarvinStocksPGM
{

    //public struct FramePoints
    //{
    //    public float frameX; 
    //    public float frameY;
    //}

    public partial class frmStocksPGM : Form
    {
        //-- 動態元件 --//
        private Label lblStokInfo;              // 動態新增label元件：顯示股票資訊用
        private Label lblHighPrice;             // 動態新增label元件：顯示股票最高價
        private Label lblLowPrice;              // 動態新增label元件：顯示股票最低價
        private Label lblHighVolume;             // 動態新增label元件：顯示成交量最高價
        private Label lblLowVolume;              // 動態新增label元件：顯示成交量最低價
        private Label lblHighBias;             // 動態新增label元件：顯示乖離率最高價
        private Label lblLowBias;              // 動態新增label元件：顯示乖離率最低價

        private Label[] lblStockYM = new Label[STOCKYM_CNTS];

        private bool _initialized = false;
        static int STOCKYM_CNTS = 36;

        private static int FrameNum = 5;            // 要顯示的frame數量
        private static int SelectFramePos = 0;      // 選擇的frame位置(1~FrameNum)
        private float XWidthBorder = 20;            // frame左、右兩邊預留的空間
        private float YHeightBorder = 10;           // frame最下面預留的空間
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
            pnlStocksBar.Width = this.Width;
            _initialized = true;
        }

        private void cboStocks_SelectedIndexChanged(object sender, EventArgs e)
        {
            StartIndex = 0;
            // 取得要顯示的股票資料
            //StkData = DbHelper.TestConnectDB(cboStocks.Text);
            StkData = StockModule.GetStockData(cboStocks.Text);
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
            this.ContextMenuStrip = cntMenuStrip;

            pnlStocksBar.Width = this.Width;
            FrameNum = int.Parse(cboFrameNum.Text);

            // 取得要顯示的股票資料
            StkData = StockModule.GetStockData(cboStocks.Text);
            IdxData = IndexModule.GetIndexData(StkData);

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
                Font = new Font(this.Font.FontFamily, 6),
            };
            this.Controls.Add(lblHighPrice);

            // 新增顯示股票資訊的標籤
            lblLowPrice = new Label()
            {
                Name = "lblLowPrice",
                Text = "low",
                AutoSize = true,
                Font = new Font(this.Font.FontFamily, 6),
            };
            this.Controls.Add(lblLowPrice);

            // 新增顯示股票資訊的標籤
            lblHighVolume = new Label()
            {
                Name = "lblHighVolume",
                Text = "high",
                AutoSize = true,
                Font = new Font(this.Font.FontFamily, 5),
            };
            this.Controls.Add(lblHighVolume);

            // 新增顯示股票資訊的標籤
            lblLowVolume = new Label()
            {
                Name = "lblLowVolume",
                Text = "low",
                AutoSize = true,
                Font = new Font(this.Font.FontFamily, 5),
            };
            this.Controls.Add(lblLowVolume);

            // 新增顯示股票資訊的標籤
            lblHighVolume = new Label()
            {
                Name = "lblHighBias",
                Text = "high",
                AutoSize = true,
                Font = new Font(this.Font.FontFamily, 5),
            };
            this.Controls.Add(lblHighVolume);

            // 新增顯示股票資訊的標籤
            lblLowVolume = new Label()
            {
                Name = "lblLowBias",
                Text = "low",
                AutoSize = true,
                Font = new Font(this.Font.FontFamily, 5),
            };
            this.Controls.Add(lblLowVolume);

            // 新增 Label 元件(預設建立 STOCKYM_CNTS 個備用)
            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i] = new Label()
                {
                    Name = $"lblStockYM{i}",
                    Text = "YYMM",
                    AutoSize = true,
                    Font = new Font(this.Font.FontFamily, 8),
                };
                this.Controls.Add(lblStockYM[i]);
            }
        }

        private void frmStocksPGM_Resize(object sender, EventArgs e)
        {
            if (!_initialized)
                return;

            pnlStocksBar.Width = this.Width;
            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i].Text = "";
            }
            this.Invalidate();
        }

        private void frmStocksPGM_Paint(object sender, PaintEventArgs e)
        {
            Debug.WriteLine($"PAINT SelectFramePos={SelectFramePos}");

            Graphics g = e.Graphics;
            e.Graphics.Clear(this.BackColor);

            // lblStokInfo : 顯示股票資訊
            float frmYTop = mnuStocksList.Size.Height + pnlStocksBar.Size.Height + lblStokInfo.Size.Height;
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

            //--------------------------------------------//
            //-- 繪製 K-Map (最上方 Frame) 的橫(虛)線段 --//
            //--------------------------------------------//
            int dashLineCounts = 5; // 虛線的段數
            float hh = FrameLeftPoints[1].frameY - FrameLeftPoints[0].frameY;

            Pen pen = new Pen(Color.Black, 1);
            pen.DashStyle = DashStyle.Dash;
            pen.DashPattern = new float[] { 4, 2 }; // 畫 4px，空 2px
            for (int i = 1; i < dashLineCounts; i++)
            {
                float x0 = FrameLeftPoints[0].frameX;
                float y0 = FrameLeftPoints[0].frameY + (hh / dashLineCounts) * i;
                float x1 = FrameMiddlePoints[0].frameX;
                float y1 = FrameMiddlePoints[0].frameY + (hh / dashLineCounts) * i;
                g.DrawLine(pen, x0, y0, x1, y1);
            }

            // 取得要顯示的股票資料
            if (StkData == null || StkData.Length == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }
            if (StartIndex > (StkData.Length - DisplayCount))
            {
                StartIndex = StkData.Length - DisplayCount;
            }
            if (StartIndex < 0)
            {
                StartIndex = 0;
            }

            // 顯示畫面筆數之最高/最低價 (因為資料庫的資料型態為 decimal，為了便於計算故宣告為 decimal)
            HighLowValues highLowValues = GeneralModule.GetHighLowValue(StkData, IdxData, StartIndex, DisplayCount, GeneralModule.MAP_K);
            lblHighPrice.Text = highLowValues.highValue.ToString("F2");
            lblLowPrice.Text = highLowValues.lowValue.ToString("F2");
            Debug.WriteLine("最高/低價：" + $"{highLowValues.highValue}, {highLowValues.lowValue}");

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
                float y0 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - StkData[i].HighPrice));
                float x1 = (barXCoord + barWidth / 2);
                float y1 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - StkData[i].LowPrice));

                if (StkData[i].StartPrice > StkData[i].EndPrice)
                    g.DrawLine(Pens.Green, x0, y0, x1, y1);
                else
                    g.DrawLine(Pens.Red, x0, y0, x1, y1);

                // 顯示股價均線 (MAP) 的線段
                if (i > StartIndex)
                {
                    Pen pen4MAP = new Pen(Color.Blue, 2);
                    // MAP5
                    float x2 = (barXCoord - barWidth + barWidth / 2);
                    float y2 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i - 1].MAP5));
                    float x3 = (barXCoord + barWidth / 2);
                    float y3 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i].MAP5));
                    g.DrawLine(pen4MAP, x2, y2, x3, y3);
                    // MAP10
                    pen4MAP = new Pen(Color.Black, 2);
                    x2 = (barXCoord - barWidth + barWidth / 2);
                    y2 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i - 1].MAP10));
                    x3 = (barXCoord + barWidth / 2);
                    y3 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i].MAP10));
                    g.DrawLine(pen4MAP, x2, y2, x3, y3);
                    // MAP20
                    pen4MAP = new Pen(Color.Orange, 2);
                    x2 = (barXCoord - barWidth + barWidth / 2);
                    y2 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i - 1].MAP20));
                    x3 = (barXCoord + barWidth / 2);
                    y3 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i].MAP20));
                    g.DrawLine(pen4MAP, x2, y2, x3, y3);
                    // MAP60
                    pen4MAP = new Pen(Color.Green, 2);
                    x2 = (barXCoord - barWidth + barWidth / 2);
                    y2 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i - 1].MAP60));
                    x3 = (barXCoord + barWidth / 2);
                    y3 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i].MAP60));
                    g.DrawLine(pen4MAP, x2, y2, x3, y3);
                    // MAP120
                    pen4MAP = new Pen(Color.Brown, 2);
                    x2 = (barXCoord - barWidth + barWidth / 2);
                    y2 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i - 1].MAP120));
                    x3 = (barXCoord + barWidth / 2);
                    y3 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i].MAP120));
                    g.DrawLine(pen4MAP, x2, y2, x3, y3);
                    // MAP240
                    pen4MAP = new Pen(Color.Violet, 2);
                    x2 = (barXCoord - barWidth + barWidth / 2);
                    y2 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i - 1].MAP240));
                    x3 = (barXCoord + barWidth / 2);
                    y3 = (float)FrameLeftPoints[0].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - IdxData[i].MAP240));
                    g.DrawLine(pen4MAP, x2, y2, x3, y3);
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
                    float y1 = FrameLeftPoints[1].frameY;
                    g.DrawLine(pen, x0, y0, x1, y1);
                    // 顯示交易日期(年月)
                    string strnum = StkData[i].TradeDate.ToString();
                    lblStockYM[k].Text = strnum.Substring(0, strnum.Length - 2);
                    lblStockYM[k].Location = new System.Drawing.Point((int)(x0 - (lblStockYM[k].Size.Width / 2)), (int)(y1 + 5));
                    k = k + 1;
                }
                prevNum = nextNum;
            }

            //=== 顯示「成交量」(MAP_VOLUME) START ===// 
            try
            {
                //Chalk_MAP_VOLUME(e.Graphics, SelectFramePos, FrameNum);
                Chalk_MAP_BIAS(e.Graphics, SelectFramePos, FrameNum);
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
            lblHighPrice.Location = new System.Drawing.Point((int)FrameLeftPoints[0].frameX - lblHighPrice.Width, (int)FrameLeftPoints[0].frameY);
            lblLowPrice.Location = new System.Drawing.Point((int)FrameLeftPoints[1].frameX - lblLowPrice.Width, (int)FrameLeftPoints[1].frameY - lblLowPrice.Height);
            //lblHighPrice.Text = highLowValues.highValue.ToString("F2");
            //lblLowPrice.Text = highLowValues.lowValue.ToString("F2");

            //=== 最右側的各項指標數值顯示 ===//
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

            //-- FOR DEBUG : Display Frame's 端點指標 --//
            for (int i = 0; i < (FrameNum + 1); i++)
            {
                Debug.WriteLine("FramePoint[" + i + "], Left.X= " + FrameLeftPoints[i].frameY + ", Left.Y= " + FrameLeftPoints[i].frameY
                                + ", Right.X= " + FrameRightPoints[i].frameY + ", Right.Y= " + FrameRightPoints[i].frameY
                                + ", Mid.X= " + FrameMiddlePoints[i].frameY + ", Mid.Y= " + FrameMiddlePoints[i].frameY);
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
            {
                this.Invalidate();
            }
        }

        private void btnBack1_Click(object sender, EventArgs e)
        {
            StartIndex += 1;

            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i].Text = "";
            }
            this.Invalidate();
        }

        private void btnFore1_Click(object sender, EventArgs e)
        {
            StartIndex -= 1;

            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i].Text = "";
            }
            this.Invalidate();
        }

        private void btnBack2_Click(object sender, EventArgs e)
        {
            StartIndex += 5;

            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i].Text = "";
            }
            this.Invalidate();
        }

        private void btnFore2_Click(object sender, EventArgs e)
        {
            StartIndex -= 5;

            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i].Text = "";
            }
            this.Invalidate();
        }

        private void btnBack3_Click(object sender, EventArgs e)
        {
            StartIndex = int.MaxValue;

            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i].Text = "";
            }
            this.Invalidate();
        }

        private void btnFore3_Click(object sender, EventArgs e)
        {
            StartIndex = 0;

            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i].Text = "";
            }
            this.Invalidate();
        }

        private void btnZoomOut_Click(object sender, EventArgs e)
        {
            DisplayCount = DisplayCount + DisplayCount / 10;

            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i].Text = "";
            }
            this.Invalidate();
        }

        private void btnZoomIn_Click(object sender, EventArgs e)
        {
            DisplayCount = DisplayCount - DisplayCount / 10;
            if (DisplayCount < 10)
                DisplayCount = 10;

            for (int i = 0; i < STOCKYM_CNTS; i++)
            {
                lblStockYM[i].Text = "";
            }
            this.Invalidate();
        }

        private void VolumeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            //Chalk_MAP_VOLUME(e.Graphics, SelectFramePos, FrameNum);
            Debug.WriteLine($"CLICK VolumeToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void BIASToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectFramePos = GetSelectFrame(FrameLeftPoints, FrameRightPoints, CursorPosition, FrameNum);
            //Chalk_MAP_BIAS(e.Graphics, SelectFramePos, FrameNum);
            Debug.WriteLine($"CLICK BIASToolStripMenuItem_Click() : SelectFramePos={SelectFramePos}");
            this.Invalidate();
        }

        private void Chalk_MAP_VOLUME(Graphics g, int framePos, int frameNum)
        {
            //========================================//
            //=== 顯示「成交量」(MAP_VOLUME) START ===// 
            //========================================//
            if (framePos <= 0 || framePos > frameNum)
                return;

            float yAxisLength = FrameLeftPoints[framePos].frameY - FrameLeftPoints[framePos - 1].frameY;    // 儲存要繪製指標柱狀圖的Y軸長度
            float barHeight = 0;                            // 要繪製指標柱狀圖高度
            float barXCoord = FrameLeftPoints[0].frameX;    // 要繪製指標柱狀圖X座標(最左邊的位置FrameLeftPoints[0]一定會存在)
            float barYCoord = 0;                            // 要繪製指標柱狀圖Y座標

            HighLowValues highLowValues = GeneralModule.GetHighLowValue(StkData, IdxData, StartIndex, DisplayCount, GeneralModule.MAP_VOLUME);
            Debug.WriteLine("最高/低價(Vol.)：" + $"{highLowValues.highValue}, {highLowValues.lowValue}");

            lblHighVolume.Text = highLowValues.highValue.ToString();
            lblLowVolume.Text = highLowValues.lowValue.ToString();
            lblHighVolume.Location = new System.Drawing.Point((int)FrameLeftPoints[framePos - 1].frameX - lblHighVolume.Width, (int)FrameLeftPoints[framePos - 1].frameY);
            lblLowVolume.Location = new System.Drawing.Point((int)FrameLeftPoints[framePos].frameX - lblLowVolume.Width, (int)FrameLeftPoints[framePos].frameY - lblLowVolume.Height);
            float yDistance = yAxisLength / (float)Math.Abs(highLowValues.highValue - highLowValues.lowValue); // 取得每個價格對應的Y軸距離

            // 劃虛線 (劃3條)
            using (Pen pen = new Pen(Color.Black, 1))
            {
                pen.DashStyle = DashStyle.Dash;
                for (int i = 1; i <= 3; i++)
                {
                    PointF pl = new PointF(FrameLeftPoints[framePos].frameX, FrameLeftPoints[framePos].frameY - yAxisLength * i / 4);
                    PointF pr = new PointF(FrameMiddlePoints[framePos].frameX, FrameMiddlePoints[framePos].frameY - yAxisLength * i / 4);
                    g.DrawLine(pen, pl, pr);
                }
            }

            for (int i = StartIndex; i < (StartIndex + DisplayCount); i++)
            {
                // X 座標
                if (i != StartIndex)
                {
                    barXCoord += FrameBarWidth;
                }
                // Y 座標
                barYCoord = (float)FrameLeftPoints[framePos - 1].frameY + (yDistance * (float)Math.Abs(highLowValues.highValue - StkData[i].Volume));
                // 計算 K-Bar 的高度
                barHeight = yDistance * (float)Math.Abs(StkData[i].Volume - highLowValues.lowValue);
                // 繪製 K-Bar
                if (StkData[i].StartPrice > StkData[i].EndPrice)
                {
                    Brush brush = new SolidBrush(Color.Green);
                    g.FillRectangle(brush, barXCoord, barYCoord, FrameBarWidth, barHeight);
                }
                else
                {
                    Brush brush = new SolidBrush(Color.Red);
                    g.FillRectangle(brush, barXCoord, barYCoord, FrameBarWidth, barHeight);
                }
            }
            Debug.WriteLine("Chalk_MAP_VOLUME() END!!!!!");
        }

        /**
         * 以BIAS為中線，向上 > 0、向下 < 0
         * g        繪圖物件
         * framePos 目前選擇的 frame 的位置
         * frameNum 目前 frame 的總數量
         */
        private void Chalk_MAP_BIAS(Graphics g, int framePos, int frameNum)
        {
            //========================================//
            //=== 顯示「乖離率」(MAP_BIAS) START   ===// 
            //========================================//
            if (framePos <= 0 || framePos > frameNum)
                return;

            float xAxisLength = FrameMiddlePoints[framePos].frameX - FrameLeftPoints[framePos].frameX;    // 儲存要繪製指標柱狀圖的X軸長度
            float yAxisHeight = FrameLeftPoints[framePos].frameY - FrameLeftPoints[framePos - 1].frameY;    // 儲存要繪製指標柱狀圖的Y軸長度
            float yDistance = yAxisHeight / 4;

            HighLowValues highLowValues = GeneralModule.GetHighLowValue(StkData, IdxData, StartIndex, DisplayCount, GeneralModule.MAP_BIAS);
            Debug.WriteLine("最高/低價(BIAS)：" + $"{highLowValues.highValue}, {highLowValues.lowValue}");

            // 劃虛線 (劃3條)
            using (Pen pen = new Pen(Color.Black, 1))
            {
                pen.DashStyle = DashStyle.Dash;
                for (int i = 1; i <= 3; i++)
                {
                    PointF pl = new PointF(FrameLeftPoints[framePos].frameX, FrameLeftPoints[framePos].frameY - yDistance * i);
                    PointF pr = new PointF(FrameMiddlePoints[framePos].frameX, FrameMiddlePoints[framePos].frameY - yDistance * i);
                    g.DrawLine(pen, pl, pr);
                }
            }

            // Draw Index Values
            PointF[] points = new PointF[DisplayCount];
            float xWidth = xAxisLength / DisplayCount;
            float yHeight = yAxisHeight / (Math.Abs((float) highLowValues.highValue) * 2f);
            Debug.WriteLine("xWidth= " + xWidth + ", yHeight= " + yHeight);

            float xCoord = FrameLeftPoints[framePos].frameX + (xWidth / 2f);
            float yCoord = FrameLeftPoints[framePos].frameY;
            Debug.WriteLine("Baseline -- framePos= " + framePos +
                ", frame[" + framePos + "].X= " + FrameLeftPoints[framePos].frameX + 
                ", frame[" + framePos + "].Y= " + FrameLeftPoints[framePos].frameY +
                ", xCoord= " + xCoord + ", yCoord= " + yCoord);

            for (int i = StartIndex; i < (StartIndex + DisplayCount); i++)
            {
                //Debug.WriteLine("i= " + i + ", StartIndex= " + StartIndex + ", DisplayCount= " + DisplayCount);
                float ii = 0;
                if (i == StartIndex)
                {
                    if (IdxData[StartIndex].BIAS >= 0)
                    {
                        ii = ((float)IdxData[StartIndex].BIAS + Math.Abs((float)highLowValues.highValue)) * yHeight;
                    }
                    else
                    {
//                        ii = ((float)highLowValues.highValue - Math.Abs((float)IdxData[StartIndex].BIAS)) * yHeight;
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
                        Debug.WriteLine("BIAS[" + i + "]>0 -- ii= " + ii + ", yDist= " + (float)highLowValues.highValue * yHeight + ", IdxData[i].BIAS= " + IdxData[i].BIAS);
                    }
                    else
                    {
                        //ii = ((float)highLowValues.highValue - Math.Abs((float)IdxData[i].BIAS)) * yHeight;
                        ii = Math.Abs((float)IdxData[i].BIAS) * yHeight;
                        Debug.WriteLine("BIAS[" + i + "]<0 -- ii= " + ii + ", yDist= " + (float)highLowValues.highValue * yHeight + ", IdxData[i].BIAS= " + IdxData[i].BIAS);
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

            // 顯示Frame最左側的標籤
            ////lblHighBias.Text = highLowValues.highValue.ToString();
            ////lblLowBias.Text = highLowValues.lowValue.ToString();
            ////lblHighBias.Location = new System.Drawing.Point((int)FrameLeftPoints[framePos - 1].frameX - lblHighBias.Width, (int)FrameLeftPoints[framePos - 1].frameY);
            ////lblLowBias.Location = new System.Drawing.Point((int)FrameLeftPoints[framePos].frameX - lblLowBias.Width, (int)FrameLeftPoints[framePos].frameY - lblLowBias.Height);
            ////float yDistance = yAxisLength / (float)Math.Abs(highLowValues.highValue - highLowValues.lowValue); // 取得每個價格對應的Y軸距離

            Debug.WriteLine("Chalk_MAP_BIAS() END!!!!!");
        }

    }
}
