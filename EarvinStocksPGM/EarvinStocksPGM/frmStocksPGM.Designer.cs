namespace EarvinStocksPGM
{
    partial class frmStocksPGM
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            mnuStocksList = new MenuStrip();
            toolStripMenuItem1 = new ToolStripMenuItem();
            日線ToolStripMenuItem = new ToolStripMenuItem();
            週線ToolStripMenuItem = new ToolStripMenuItem();
            月線ToolStripMenuItem = new ToolStripMenuItem();
            選項OToolStripMenuItem = new ToolStripMenuItem();
            移動查價ToolStripMenuItem = new ToolStripMenuItem();
            設定指標天數ToolStripMenuItem = new ToolStripMenuItem();
            列印畫面ToolStripMenuItem = new ToolStripMenuItem();
            大小VToolStripMenuItem = new ToolStripMenuItem();
            放大ToolStripMenuItem = new ToolStripMenuItem();
            縮小ToolStripMenuItem = new ToolStripMenuItem();
            說明HToolStripMenuItem = new ToolStripMenuItem();
            關於ToolStripMenuItem = new ToolStripMenuItem();
            pnlStocksBar = new Panel();
            cboStocks = new ComboBox();
            cboStocksType = new ComboBox();
            cboStocksFrom = new ComboBox();
            label3 = new Label();
            cboFrameNum = new ComboBox();
            label2 = new Label();
            btnBack3 = new Button();
            btnBack2 = new Button();
            btnBack1 = new Button();
            btnFore1 = new Button();
            btnFore2 = new Button();
            btnFocus = new Button();
            btnFore3 = new Button();
            btnZoomOut = new Button();
            btnZoomIn = new Button();
            label1 = new Label();
            mnuStocksList.SuspendLayout();
            pnlStocksBar.SuspendLayout();
            SuspendLayout();
            // 
            // mnuStocksList
            // 
            mnuStocksList.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, 選項OToolStripMenuItem, 大小VToolStripMenuItem, 說明HToolStripMenuItem });
            mnuStocksList.Location = new Point(0, 0);
            mnuStocksList.Name = "mnuStocksList";
            mnuStocksList.Size = new Size(1184, 24);
            mnuStocksList.TabIndex = 1;
            mnuStocksList.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.DropDownItems.AddRange(new ToolStripItem[] { 日線ToolStripMenuItem, 週線ToolStripMenuItem, 月線ToolStripMenuItem });
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(55, 20);
            toolStripMenuItem1.Text = "K線(D)";
            toolStripMenuItem1.Click += toolStripMenuItem1_Click;
            // 
            // 日線ToolStripMenuItem
            // 
            日線ToolStripMenuItem.Name = "日線ToolStripMenuItem";
            日線ToolStripMenuItem.Size = new Size(98, 22);
            日線ToolStripMenuItem.Text = "日線";
            // 
            // 週線ToolStripMenuItem
            // 
            週線ToolStripMenuItem.Name = "週線ToolStripMenuItem";
            週線ToolStripMenuItem.Size = new Size(98, 22);
            週線ToolStripMenuItem.Text = "週線";
            // 
            // 月線ToolStripMenuItem
            // 
            月線ToolStripMenuItem.Name = "月線ToolStripMenuItem";
            月線ToolStripMenuItem.Size = new Size(98, 22);
            月線ToolStripMenuItem.Text = "月線";
            // 
            // 選項OToolStripMenuItem
            // 
            選項OToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 移動查價ToolStripMenuItem, 設定指標天數ToolStripMenuItem, 列印畫面ToolStripMenuItem });
            選項OToolStripMenuItem.Name = "選項OToolStripMenuItem";
            選項OToolStripMenuItem.Size = new Size(61, 20);
            選項OToolStripMenuItem.Text = "選項(O)";
            // 
            // 移動查價ToolStripMenuItem
            // 
            移動查價ToolStripMenuItem.Name = "移動查價ToolStripMenuItem";
            移動查價ToolStripMenuItem.Size = new Size(146, 22);
            移動查價ToolStripMenuItem.Text = "移動查價";
            // 
            // 設定指標天數ToolStripMenuItem
            // 
            設定指標天數ToolStripMenuItem.Name = "設定指標天數ToolStripMenuItem";
            設定指標天數ToolStripMenuItem.Size = new Size(146, 22);
            設定指標天數ToolStripMenuItem.Text = "設定指標天數";
            // 
            // 列印畫面ToolStripMenuItem
            // 
            列印畫面ToolStripMenuItem.Name = "列印畫面ToolStripMenuItem";
            列印畫面ToolStripMenuItem.Size = new Size(146, 22);
            列印畫面ToolStripMenuItem.Text = "列印畫面";
            // 
            // 大小VToolStripMenuItem
            // 
            大小VToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 放大ToolStripMenuItem, 縮小ToolStripMenuItem });
            大小VToolStripMenuItem.Name = "大小VToolStripMenuItem";
            大小VToolStripMenuItem.Size = new Size(59, 20);
            大小VToolStripMenuItem.Text = "大小(V)";
            // 
            // 放大ToolStripMenuItem
            // 
            放大ToolStripMenuItem.Name = "放大ToolStripMenuItem";
            放大ToolStripMenuItem.Size = new Size(98, 22);
            放大ToolStripMenuItem.Text = "放大";
            // 
            // 縮小ToolStripMenuItem
            // 
            縮小ToolStripMenuItem.Name = "縮小ToolStripMenuItem";
            縮小ToolStripMenuItem.Size = new Size(98, 22);
            縮小ToolStripMenuItem.Text = "縮小";
            // 
            // 說明HToolStripMenuItem
            // 
            說明HToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 關於ToolStripMenuItem });
            說明HToolStripMenuItem.Name = "說明HToolStripMenuItem";
            說明HToolStripMenuItem.Size = new Size(60, 20);
            說明HToolStripMenuItem.Text = "說明(H)";
            // 
            // 關於ToolStripMenuItem
            // 
            關於ToolStripMenuItem.Name = "關於ToolStripMenuItem";
            關於ToolStripMenuItem.Size = new Size(98, 22);
            關於ToolStripMenuItem.Text = "關於";
            // 
            // pnlStocksBar
            // 
            pnlStocksBar.BackColor = SystemColors.ActiveCaption;
            pnlStocksBar.Controls.Add(cboStocks);
            pnlStocksBar.Controls.Add(cboStocksType);
            pnlStocksBar.Controls.Add(cboStocksFrom);
            pnlStocksBar.Controls.Add(label3);
            pnlStocksBar.Controls.Add(cboFrameNum);
            pnlStocksBar.Controls.Add(label2);
            pnlStocksBar.Controls.Add(btnBack3);
            pnlStocksBar.Controls.Add(btnBack2);
            pnlStocksBar.Controls.Add(btnBack1);
            pnlStocksBar.Controls.Add(btnFore1);
            pnlStocksBar.Controls.Add(btnFore2);
            pnlStocksBar.Controls.Add(btnFocus);
            pnlStocksBar.Controls.Add(btnFore3);
            pnlStocksBar.Controls.Add(btnZoomOut);
            pnlStocksBar.Controls.Add(btnZoomIn);
            pnlStocksBar.Controls.Add(label1);
            pnlStocksBar.Location = new Point(0, 27);
            pnlStocksBar.Name = "pnlStocksBar";
            pnlStocksBar.Size = new Size(1008, 35);
            pnlStocksBar.TabIndex = 2;
            pnlStocksBar.Paint += pnlStocksBar_Paint;
            // 
            // cboStocks
            // 
            cboStocks.FormattingEnabled = true;
            cboStocks.Location = new Point(726, 4);
            cboStocks.Name = "cboStocks";
            cboStocks.Size = new Size(121, 28);
            cboStocks.TabIndex = 15;
            cboStocks.SelectedIndexChanged += cboStocks_SelectedIndexChanged;
            // 
            // cboStocksType
            // 
            cboStocksType.FormattingEnabled = true;
            cboStocksType.Items.AddRange(new object[] { "日線", "週線", "月線" });
            cboStocksType.Location = new Point(575, 5);
            cboStocksType.Name = "cboStocksType";
            cboStocksType.Size = new Size(70, 28);
            cboStocksType.TabIndex = 14;
            cboStocksType.Text = "日線";
            // 
            // cboStocksFrom
            // 
            cboStocksFrom.FormattingEnabled = true;
            cboStocksFrom.Items.AddRange(new object[] { "File", "Directory" });
            cboStocksFrom.Location = new Point(486, 5);
            cboStocksFrom.Name = "cboStocksFrom";
            cboStocksFrom.Size = new Size(85, 28);
            cboStocksFrom.TabIndex = 13;
            cboStocksFrom.Text = "File";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(651, 8);
            label3.Name = "label3";
            label3.Size = new Size(73, 20);
            label3.TabIndex = 12;
            label3.Text = "選擇個股";
            // 
            // cboFrameNum
            // 
            cboFrameNum.FormattingEnabled = true;
            cboFrameNum.Items.AddRange(new object[] { "2", "3", "4", "5", "6", "7", "8", "9" });
            cboFrameNum.Location = new Point(67, 5);
            cboFrameNum.Name = "cboFrameNum";
            cboFrameNum.Size = new Size(42, 28);
            cboFrameNum.TabIndex = 11;
            cboFrameNum.Text = "5";
            cboFrameNum.SelectedIndexChanged += cboFrameNum_SelectedIndexChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(415, 9);
            label2.Name = "label2";
            label2.Size = new Size(73, 20);
            label2.TabIndex = 10;
            label2.Text = "股票來源";
            label2.Click += label2_Click;
            // 
            // btnBack3
            // 
            btnBack3.FlatStyle = FlatStyle.Popup;
            btnBack3.Image = EarvinStocksPGM.Properties.Resources.BACK3;
            btnBack3.ImageAlign = ContentAlignment.TopLeft;
            btnBack3.Location = new Point(298, 8);
            btnBack3.Name = "btnBack3";
            btnBack3.Size = new Size(30, 25);
            btnBack3.TabIndex = 9;
            btnBack3.UseVisualStyleBackColor = true;
            // 
            // btnBack2
            // 
            btnBack2.FlatStyle = FlatStyle.Popup;
            btnBack2.Image = EarvinStocksPGM.Properties.Resources.BACK2;
            btnBack2.ImageAlign = ContentAlignment.TopLeft;
            btnBack2.Location = new Point(271, 8);
            btnBack2.Name = "btnBack2";
            btnBack2.Size = new Size(30, 25);
            btnBack2.TabIndex = 8;
            btnBack2.UseVisualStyleBackColor = true;
            btnBack2.Click += btnBack2_Click;
            // 
            // btnBack1
            // 
            btnBack1.FlatStyle = FlatStyle.Popup;
            btnBack1.Image = EarvinStocksPGM.Properties.Resources.BACK1;
            btnBack1.ImageAlign = ContentAlignment.TopLeft;
            btnBack1.Location = new Point(244, 8);
            btnBack1.Name = "btnBack1";
            btnBack1.Size = new Size(30, 25);
            btnBack1.TabIndex = 7;
            btnBack1.UseVisualStyleBackColor = true;
            // 
            // btnFore1
            // 
            btnFore1.FlatStyle = FlatStyle.Popup;
            btnFore1.Image = EarvinStocksPGM.Properties.Resources.FORE1;
            btnFore1.ImageAlign = ContentAlignment.TopLeft;
            btnFore1.Location = new Point(218, 8);
            btnFore1.Name = "btnFore1";
            btnFore1.Size = new Size(30, 25);
            btnFore1.TabIndex = 6;
            btnFore1.UseVisualStyleBackColor = true;
            // 
            // btnFore2
            // 
            btnFore2.FlatStyle = FlatStyle.Popup;
            btnFore2.Image = EarvinStocksPGM.Properties.Resources.FORE2;
            btnFore2.ImageAlign = ContentAlignment.TopLeft;
            btnFore2.Location = new Point(192, 8);
            btnFore2.Name = "btnFore2";
            btnFore2.Size = new Size(30, 25);
            btnFore2.TabIndex = 5;
            btnFore2.UseVisualStyleBackColor = true;
            // 
            // btnFocus
            // 
            btnFocus.BackColor = SystemColors.ActiveBorder;
            btnFocus.FlatStyle = FlatStyle.Popup;
            btnFocus.Location = new Point(334, 4);
            btnFocus.Name = "btnFocus";
            btnFocus.Size = new Size(75, 30);
            btnFocus.TabIndex = 4;
            btnFocus.Text = "查價";
            btnFocus.UseVisualStyleBackColor = false;
            btnFocus.Click += btnFocus_Click;
            // 
            // btnFore3
            // 
            btnFore3.FlatStyle = FlatStyle.Popup;
            btnFore3.Image = EarvinStocksPGM.Properties.Resources.FORE3;
            btnFore3.ImageAlign = ContentAlignment.TopLeft;
            btnFore3.Location = new Point(166, 8);
            btnFore3.Name = "btnFore3";
            btnFore3.Size = new Size(30, 25);
            btnFore3.TabIndex = 3;
            btnFore3.UseVisualStyleBackColor = true;
            // 
            // btnZoomOut
            // 
            btnZoomOut.FlatStyle = FlatStyle.Popup;
            btnZoomOut.Image = EarvinStocksPGM.Properties.Resources.ZOOMOUT;
            btnZoomOut.Location = new Point(140, 8);
            btnZoomOut.Name = "btnZoomOut";
            btnZoomOut.Size = new Size(30, 25);
            btnZoomOut.TabIndex = 2;
            btnZoomOut.UseVisualStyleBackColor = true;
            // 
            // btnZoomIn
            // 
            btnZoomIn.FlatStyle = FlatStyle.Popup;
            btnZoomIn.Image = EarvinStocksPGM.Properties.Resources.ZOOMIN;
            btnZoomIn.Location = new Point(115, 8);
            btnZoomIn.Name = "btnZoomIn";
            btnZoomIn.Size = new Size(30, 25);
            btnZoomIn.TabIndex = 1;
            btnZoomIn.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(4, 8);
            label1.Name = "label1";
            label1.Size = new Size(57, 20);
            label1.TabIndex = 0;
            label1.Text = "視窗數";
            // 
            // frmStocksPGM
            // 
            AutoScaleDimensions = new SizeF(10F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1184, 761);
            Controls.Add(pnlStocksBar);
            Controls.Add(mnuStocksList);
            Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 136);
            MainMenuStrip = mnuStocksList;
            Margin = new Padding(4);
            Name = "frmStocksPGM";
            Text = "Stocks Test Form";
            Load += frmStocksPGM_Load;
            Paint += frmStocksPGM_Paint;
            Resize += frmStocksPGM_Resize;
            mnuStocksList.ResumeLayout(false);
            mnuStocksList.PerformLayout();
            pnlStocksBar.ResumeLayout(false);
            pnlStocksBar.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private MenuStrip mnuStocksList;
        private ToolStripMenuItem toolStripMenuItem1;
        private ToolStripMenuItem 日線ToolStripMenuItem;
        private ToolStripMenuItem 週線ToolStripMenuItem;
        private ToolStripMenuItem 月線ToolStripMenuItem;
        private ToolStripMenuItem 選項OToolStripMenuItem;
        private ToolStripMenuItem 移動查價ToolStripMenuItem;
        private ToolStripMenuItem 設定指標天數ToolStripMenuItem;
        private ToolStripMenuItem 列印畫面ToolStripMenuItem;
        private ToolStripMenuItem 大小VToolStripMenuItem;
        private ToolStripMenuItem 放大ToolStripMenuItem;
        private ToolStripMenuItem 縮小ToolStripMenuItem;
        private ToolStripMenuItem 說明HToolStripMenuItem;
        private ToolStripMenuItem 關於ToolStripMenuItem;
        private Panel pnlStocksBar;
        private Label label1;
        private Button btnFocus;
        private Button btnFore3;
        private Button btnZoomOut;
        private Button btnZoomIn;
        private Button btnBack1;
        private Button btnFore1;
        private Button btnFore2;
        private Button btnBack3;
        private Button btnBack2;
        private ComboBox cboFrameNum;
        private Label label2;
        private Label label3;
        private ComboBox cboStocksFrom;
        private ComboBox cboStocksType;
        private ComboBox cboStocks;
    }
}
