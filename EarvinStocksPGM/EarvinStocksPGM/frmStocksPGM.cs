using System.Diagnostics;
using MySqlConnector;



namespace EarvinStocksPGM
{
    struct FramePoints
    {
        public int frameX;
        public int frameY;
    }



    public partial class frmStocksPGM : Form
    {
        private int XWidthBorder = 20;      // frame左、右兩邊預留的空間
        private int YHeightBorder = 10;     // frame最下面預留的空間
        private int frmXTop = 30;           // frame最左上角的X座標
        private int frmRightBorder = 150;   // frame最左上角的Y座標

        private Label lblStokInfo;          // 動態新增label元件：顯示股票資訊用

        public frmStocksPGM()
        {
            InitializeComponent();

            this.StartPosition = FormStartPosition.CenterScreen;
            pnlStocksBar.Width = this.Width;
        }


        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void pnlStocksBar_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnBack2_Click(object sender, EventArgs e)
        {

        }

        private void btnFocus_Click(object sender, EventArgs e)
        {
            //            DbHelper.StockData[] sd = DbHelper.TestConnectDB();
            DbHelper.StockData[] sd = DbHelper.TestConnectDB();
            if (sd == null || sd.Length == 0)      
            {    
                MessageBox.Show("沒有資料");
                return;
            }
            
            for (int j = 0; j < sd.Length; j++)
            {
                Debug.WriteLine($"{sd[j].TradeDate}, {sd[j].StartPrice}, {sd[j].EndPrice}");
            }

        }

        private void frmStocksPGM_Load(object sender, EventArgs e)
        {
            pnlStocksBar.Width = this.Width;

            // 新增顯示股票資訊的標籤
            lblStokInfo = new Label()
            {
                Name = "lblStokInfo",
                Text = "This is a test message!",
                AutoSize = true,
                //                Size = new System.Drawing.Size(200, 40),
                Location = new Point(10, mnuStocksList.Size.Height + pnlStocksBar.Size.Height)

            };
            this.Controls.Add(lblStokInfo);
            // 若要使用名為 Home 的資源，請在專案資源中新增該影像，或改用現有資源名稱 (例如 ZOOMIN)
        }

        private void cboStocks_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void frmStocksPGM_Resize(object sender, EventArgs e)
        {
            pnlStocksBar.Width = this.Width;
            //            pnlStocksBar.Height = this.Height;
            this.Invalidate();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void frmStocksPGM_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            e.Graphics.Clear(this.BackColor);

            //int XWidthBorder = 20;
            //int YHeightBorder = 10;
            //int frmXTop = 10;

            // 40 : 要示在pnlStocksBar 與 frame間要預留26個pixels顯示股票資訊
            int frmYTop = mnuStocksList.Size.Height + pnlStocksBar.Size.Height + lblStokInfo.Size.Height;
//            int frmYTop = mnuStocksList.Size.Height + pnlStocksBar.Size.Height + 40;
            // XWidthBorder : 表示frame左右皆各內縮 (XWidthBorder / 2) 個pixels
            int frmXWidth = this.ClientSize.Width - frmXTop - (XWidthBorder / 2);
            // YHeightBorder : 表示frame最下面上調YHeightBorder個pixels
            int frmYHeight = this.ClientSize.Height - frmYTop - YHeightBorder;

            int frameNum = int.Parse(cboFrameNum.Text);
            Debug.WriteLine("frameNum= " + frameNum);

            FramePoints[] frameLeftPoints = new FramePoints[frameNum];
            FramePoints[] frameRightPoints = new FramePoints[frameNum];

            for (int i = 0; i < frameNum; i++)
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
            }

            for (int i = 0; i < frameNum; i++)
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
            }

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
            }
            g.DrawLine(Pens.Magenta, frameMiddlePoints[0].frameX, frameMiddlePoints[0].frameY, frameMiddlePoints[frameNum].frameX, frameMiddlePoints[frameNum].frameY);
            g.DrawRectangle(Pens.Blue, frmXTop, frmYTop, frmXWidth, frmYHeight);
            for (int i = 1; i < frameNum; i++)
            {
                g.DrawLine(Pens.Brown, frameLeftPoints[i].frameX, frameLeftPoints[i].frameY, frameRightPoints[i].frameX, frameRightPoints[i].frameY);
            }

            //// 顯示frame 各個點的座標
            //for (int i = 0; i < frameNum; i++)
            //{
            //    Debug.WriteLine("左: " + $"frameLeftPoints[{i}]: ({frameLeftPoints[i].frameX}, {frameLeftPoints[i].frameY})");
            //    Debug.WriteLine("右: " + $"frameRightPoints[{i}]: ({frameRightPoints[i].frameX}, {frameRightPoints[i].frameY})");
            //    Debug.WriteLine("中: " + $"frameMiddlePoints[{i}]: ({frameMiddlePoints[i].frameX}, {frameMiddlePoints[i].frameY})");
            //}
            //Debug.WriteLine("中: " + $"frameMiddlePoints[{frameNum}]: ({frameMiddlePoints[frameNum].frameX}, {frameMiddlePoints[frameNum].frameY})");


            int displayCount = 30; // 顯示的資料筆數
            decimal stockPriceHighest = 0;
            decimal stockProceLowest = 99999;

            //-----------------------------------------------
            // 取得要顯示的股票資料
            //-----------------------------------------------
            DbHelper.StockData[] sd = DbHelper.TestConnectDB();
            if (sd == null || sd.Length == 0)
            {
                MessageBox.Show("沒有資料");
                return;
            }

//            for (int j = 0; j < sd.Length; j++)
              for (int j = 0; j < 30; j++)
                {
                    if (stockPriceHighest < sd[j].EndPrice)
                    stockPriceHighest = sd[j].EndPrice;
                if (stockProceLowest > sd[j].EndPrice)
                    stockProceLowest = sd[j].EndPrice;
            }
            Debug.WriteLine("最高/低價：" + $"{stockPriceHighest}, {stockProceLowest}");

            // X-Length : 顯示frame的X軸長度；Y-Length : 顯示frame的Y軸長度
            int Xlength = frameMiddlePoints[0].frameX - frameLeftPoints[0].frameX;
            int Ylength = frameLeftPoints[1].frameY - frameLeftPoints[0].frameY;

            Debug.WriteLine("x-length: " + Xlength + " ; y-length: " + Ylength);

            float barWidth = Xlength / 30;
            float barHeight = 0;
            float barX = frameLeftPoints[0].frameX;
            float barY = 0;

            float yDistance = (float)Ylength / (float)Math.Abs(stockPriceHighest - stockProceLowest);

            for (int i = 0; i < 30; i++)
            {
                if (i !=0)
                {
                    barX += barWidth;
                }
//                barY = (float)frameLeftPoints[0].frameY + (float)Ylength / (float)Math.Abs(stockPriceHighest - stockProceLowest) * (float)Math.Abs(sd[i].StartPrice - sd[i].EndPrice);
                barY = (float)frameLeftPoints[0].frameY + yDistance * (float)Math.Abs(sd[i].StartPrice - sd[i].EndPrice);
                barHeight = (float)Ylength / (float)Math.Abs(stockPriceHighest - stockProceLowest) * (float)Math.Abs(sd[i].StartPrice - sd[i].EndPrice);

                Debug.WriteLine("aaa= " + barX + "\t\t," + barY + "\t\t," + barWidth + "\t\t," + barHeight);
                if (barHeight != 0)
                    g.DrawRectangle(Pens.Black, barX, barY, barWidth, barHeight);
                //
                float x0 = (barX + barWidth / 2);
//                float y0 = ((float)frameLeftPoints[0].frameY + yDistance * (float)Math.Abs(stockPriceHighest - sd[i].HighPrice));
                float y0 = (float)frameLeftPoints[0].frameY + (Ylength - yDistance * (float)Math.Abs(stockPriceHighest - sd[i].HighPrice));
                float x1 = (barX + barWidth / 2);
//                float y1 (float)frameLeftPoints[0].frameY= ((float)frameLeftPoints[0].frameY + yDistance * (float)Math.Abs(stockPriceHighest - sd[i].LowPrice));
                float y1 = (float)frameLeftPoints[0].frameY + (Ylength - yDistance * (float)Math.Abs(stockPriceHighest - sd[i].LowPrice));
                Debug.WriteLine(">>> x0: " + x0 + " ,y0: " + y0, " ,x1: " + x1 + " ,y1: " + y1);
                g.DrawLine(Pens.Red, x0, y0, x1, y1);
            }

            // Y-Length : 顯示frame的Y軸長度

            lblStokInfo.Text = "股票資訊";
        }

        private void cboFrameNum_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Invalidate();
        }
    }
}
