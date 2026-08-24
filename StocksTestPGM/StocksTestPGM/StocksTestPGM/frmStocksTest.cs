using System.Diagnostics;
using MySqlConnector;

namespace StocksTestPGM
{

    struct FramePoints
    {
        public int frameX;
        public int frameY;
    }



    public partial class frmStocksTest : Form
    {
        public frmStocksTest()
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

            string connStr =
                "Server=localhost;Database=stocksdb;User ID=root;Password=lin32ledi;";

            using var conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();
                Console.WriteLine("連線成功");
                MessageBox.Show("連線成功");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"連線失敗: {ex.Message}");
                MessageBox.Show("連線失敗");
            }
        }

        private void frmStocksTest_Load(object sender, EventArgs e)
        {
            pnlStocksBar.Width = this.Width;

            // 若要使用名為 Home 的資源，請在專案資源中新增該影像，或改用現有資源名稱 (例如 ZOOMIN)
        }

        private void cboStocks_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void frmStocksTest_Resize(object sender, EventArgs e)
        {
            pnlStocksBar.Width = this.Width;
            //            pnlStocksBar.Height = this.Height;
            this.Invalidate();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void frmStocksTest_Paint(object sender, PaintEventArgs e)
        {
            int frmXTop = 10;
            int frmYTop = 24 + 35 + 26;
            int frmXWidth = this.ClientSize.Width - 20;
            int frmYHeight = this.ClientSize.Height - 85 - 10;

            Graphics g = e.Graphics;
            e.Graphics.Clear(this.BackColor);

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
                frameMiddlePoints[i].frameX = frmXTop + (frmXWidth - 150);
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
            //frameMiddlePoints[0].frameX = frmXTop + (frmXWidth - 150);
            //frameMiddlePoints[0].frameY = frmYTop;
            //frameMiddlePoints[1].frameX = frmXTop + (frmXWidth - 150);
            //frameMiddlePoints[1].frameY = frmYTop + frmYHeight;
            //g.DrawLine(Pens.Magenta, frameMiddlePoints[0].frameX, frameMiddlePoints[0].frameY, frameMiddlePoints[1].frameX, frameMiddlePoints[1].frameY);

            g.DrawRectangle(Pens.Blue, frmXTop, frmYTop, frmXWidth, frmYHeight);
            for (int i = 1; i < frameNum; i++)
            {
                g.DrawLine(Pens.Brown, frameLeftPoints[i].frameX, frameLeftPoints[i].frameY, frameRightPoints[i].frameX, frameRightPoints[i].frameY);
            }
            // 
        }

        private void cboFrameNum_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Invalidate();
        }
    }
}
