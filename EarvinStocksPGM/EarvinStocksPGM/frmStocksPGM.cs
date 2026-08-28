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
        private int XWidthBorder = 20;
        private int YHeightBorder = 10;
        private int frmXTop = 10;
        private int frmRightBorder = 150;

        private Label lblStokInfo;

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
            //MessageBox.Show($"寬度: {this.Width}, 高度: {this.Height}");
            //MessageBox.Show($"寬度: {this.ClientSize.Width}, 高度: {this.ClientSize.Height}");

            //string connStr =
            //    "Server=localhost;Database=stocksdb;User ID=root;Password=lin32ledi;";

            //using var conn = new MySqlConnection(connStr);

            //try
            //{
            //    conn.Open();
            //    Console.WriteLine("連線成功");
            //    MessageBox.Show("連線成功");
            //    string sql = "SELECT DATE, START_PRICE, HIGH_PRICE, LOW_PRICE, END_PRICE, VOLUME FROM TAIWAN_DATA_POLARIS WHERE STOCK_NO = @stock_no ORDER BY DATE ";
            //    MySqlCommand cmd = new MySqlCommand(sql, conn);
            //    cmd.Parameters.AddWithValue("@stock_no", "1101");
            //    MySqlDataReader reader = cmd.ExecuteReader();
            //    while (reader.Read())
            //    {
            //        System.Diagnostics.Debug.WriteLine($"{reader["DATE"]}, " + $"{reader["START_PRICE"]}, " + $"{reader["END_PRICE"]}");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"連線失敗: {ex.Message}");
            //    MessageBox.Show("連線失敗");
            //}
            DbHelper.TestConnectDB();
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
            int frmXWidth = this.ClientSize.Width - XWidthBorder;
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
        }

        private void cboFrameNum_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Invalidate();
        }
    }
}
