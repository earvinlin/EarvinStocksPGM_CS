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
            components = new System.ComponentModel.Container();
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
            lblMAP1 = new Label();
            lblMAP2 = new Label();
            lblMAP3 = new Label();
            lblMAP4 = new Label();
            lblMAP5 = new Label();
            lblMAP6 = new Label();
            lblMAV1 = new Label();
            lblMAV2 = new Label();
            lblMAV3 = new Label();
            lblMAV4 = new Label();
            lblMAV5 = new Label();
            cntMenuStrip = new ContextMenuStrip(components);
            k線ToolStripMenuItem = new ToolStripMenuItem();
            移動平均線ToolStripMenuItem = new ToolStripMenuItem();
            量能指標ToolStripMenuItem = new ToolStripMenuItem();
            VolumeToolStripMenuItem = new ToolStripMenuItem();
            融資餘額ToolStripMenuItem = new ToolStripMenuItem();
            融資增減ToolStripMenuItem = new ToolStripMenuItem();
            融券餘額ToolStripMenuItem = new ToolStripMenuItem();
            法人庫存ToolStripMenuItem = new ToolStripMenuItem();
            自營商庫存ToolStripMenuItem = new ToolStripMenuItem();
            投信庫存ToolStripMenuItem = new ToolStripMenuItem();
            外資庫存ToolStripMenuItem = new ToolStripMenuItem();
            熱門指標ToolStripMenuItem = new ToolStripMenuItem();
            KDToolStripMenuItem = new ToolStripMenuItem();
            MACDToolStripMenuItem = new ToolStripMenuItem();
            RSIToolStripMenuItem = new ToolStripMenuItem();
            stochRSIToolStripMenuItem = new ToolStripMenuItem();
            wildersRSIToolStripMenuItem = new ToolStripMenuItem();
            WMSToolStripMenuItem = new ToolStripMenuItem();
            WMSCToolStripMenuItem = new ToolStripMenuItem();
            EMAToolStripMenuItem = new ToolStripMenuItem();
            BIASToolStripMenuItem = new ToolStripMenuItem();
            dFTSToolStripMenuItem = new ToolStripMenuItem();
            trendToolStripMenuItem = new ToolStripMenuItem();
            qMToolStripMenuItem = new ToolStripMenuItem();
            wLSTToolStripMenuItem = new ToolStripMenuItem();
            wCYToolStripMenuItem = new ToolStripMenuItem();
            wLWToolStripMenuItem = new ToolStripMenuItem();
            dIFFToolStripMenuItem = new ToolStripMenuItem();
            dCToolStripMenuItem = new ToolStripMenuItem();
            holdToolStripMenuItem = new ToolStripMenuItem();
            profitToolStripMenuItem = new ToolStripMenuItem();
            cYToolStripMenuItem = new ToolStripMenuItem();
            signalsToolStripMenuItem = new ToolStripMenuItem();
            灰聚類模型ToolStripMenuItem = new ToolStripMenuItem();
            切割區間ToolStripMenuItem = new ToolStripMenuItem();
            聚類結果ToolStripMenuItem = new ToolStripMenuItem();
            mnuStocksList.SuspendLayout();
            pnlStocksBar.SuspendLayout();
            cntMenuStrip.SuspendLayout();
            SuspendLayout();
            // 
            // mnuStocksList
            // 
            mnuStocksList.ImageScalingSize = new Size(24, 24);
            mnuStocksList.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, 選項OToolStripMenuItem, 大小VToolStripMenuItem, 說明HToolStripMenuItem });
            mnuStocksList.Location = new Point(0, 0);
            mnuStocksList.Name = "mnuStocksList";
            mnuStocksList.Size = new Size(1184, 31);
            mnuStocksList.TabIndex = 1;
            mnuStocksList.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.DropDownItems.AddRange(new ToolStripItem[] { 日線ToolStripMenuItem, 週線ToolStripMenuItem, 月線ToolStripMenuItem });
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(81, 27);
            toolStripMenuItem1.Text = "K線(D)";
            // 
            // 日線ToolStripMenuItem
            // 
            日線ToolStripMenuItem.Name = "日線ToolStripMenuItem";
            日線ToolStripMenuItem.Size = new Size(146, 34);
            日線ToolStripMenuItem.Text = "日線";
            // 
            // 週線ToolStripMenuItem
            // 
            週線ToolStripMenuItem.Name = "週線ToolStripMenuItem";
            週線ToolStripMenuItem.Size = new Size(146, 34);
            週線ToolStripMenuItem.Text = "週線";
            // 
            // 月線ToolStripMenuItem
            // 
            月線ToolStripMenuItem.Name = "月線ToolStripMenuItem";
            月線ToolStripMenuItem.Size = new Size(146, 34);
            月線ToolStripMenuItem.Text = "月線";
            // 
            // 選項OToolStripMenuItem
            // 
            選項OToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 移動查價ToolStripMenuItem, 設定指標天數ToolStripMenuItem, 列印畫面ToolStripMenuItem });
            選項OToolStripMenuItem.Name = "選項OToolStripMenuItem";
            選項OToolStripMenuItem.Size = new Size(89, 27);
            選項OToolStripMenuItem.Text = "選項(O)";
            // 
            // 移動查價ToolStripMenuItem
            // 
            移動查價ToolStripMenuItem.Name = "移動查價ToolStripMenuItem";
            移動查價ToolStripMenuItem.Size = new Size(218, 34);
            移動查價ToolStripMenuItem.Text = "移動查價";
            // 
            // 設定指標天數ToolStripMenuItem
            // 
            設定指標天數ToolStripMenuItem.Name = "設定指標天數ToolStripMenuItem";
            設定指標天數ToolStripMenuItem.Size = new Size(218, 34);
            設定指標天數ToolStripMenuItem.Text = "設定指標天數";
            // 
            // 列印畫面ToolStripMenuItem
            // 
            列印畫面ToolStripMenuItem.Name = "列印畫面ToolStripMenuItem";
            列印畫面ToolStripMenuItem.Size = new Size(218, 34);
            列印畫面ToolStripMenuItem.Text = "列印畫面";
            // 
            // 大小VToolStripMenuItem
            // 
            大小VToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 放大ToolStripMenuItem, 縮小ToolStripMenuItem });
            大小VToolStripMenuItem.Name = "大小VToolStripMenuItem";
            大小VToolStripMenuItem.Size = new Size(86, 27);
            大小VToolStripMenuItem.Text = "大小(V)";
            // 
            // 放大ToolStripMenuItem
            // 
            放大ToolStripMenuItem.Name = "放大ToolStripMenuItem";
            放大ToolStripMenuItem.Size = new Size(146, 34);
            放大ToolStripMenuItem.Text = "放大";
            // 
            // 縮小ToolStripMenuItem
            // 
            縮小ToolStripMenuItem.Name = "縮小ToolStripMenuItem";
            縮小ToolStripMenuItem.Size = new Size(146, 34);
            縮小ToolStripMenuItem.Text = "縮小";
            // 
            // 說明HToolStripMenuItem
            // 
            說明HToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 關於ToolStripMenuItem });
            說明HToolStripMenuItem.Name = "說明HToolStripMenuItem";
            說明HToolStripMenuItem.Size = new Size(88, 27);
            說明HToolStripMenuItem.Text = "說明(H)";
            // 
            // 關於ToolStripMenuItem
            // 
            關於ToolStripMenuItem.Name = "關於ToolStripMenuItem";
            關於ToolStripMenuItem.Size = new Size(146, 34);
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
            // 
            // cboStocks
            // 
            cboStocks.FormattingEnabled = true;
            cboStocks.Items.AddRange(new object[] { "1101", "2002", "00878" });
            cboStocks.Location = new Point(726, 4);
            cboStocks.Name = "cboStocks";
            cboStocks.Size = new Size(121, 38);
            cboStocks.TabIndex = 15;
            cboStocks.Text = "1101";
            cboStocks.SelectedIndexChanged += cboStocks_SelectedIndexChanged;
            // 
            // cboStocksType
            // 
            cboStocksType.FormattingEnabled = true;
            cboStocksType.Items.AddRange(new object[] { "日線", "週線", "月線" });
            cboStocksType.Location = new Point(575, 5);
            cboStocksType.Name = "cboStocksType";
            cboStocksType.Size = new Size(70, 38);
            cboStocksType.TabIndex = 14;
            cboStocksType.Text = "日線";
            // 
            // cboStocksFrom
            // 
            cboStocksFrom.FormattingEnabled = true;
            cboStocksFrom.Items.AddRange(new object[] { "File", "Directory" });
            cboStocksFrom.Location = new Point(486, 5);
            cboStocksFrom.Name = "cboStocksFrom";
            cboStocksFrom.Size = new Size(85, 38);
            cboStocksFrom.TabIndex = 13;
            cboStocksFrom.Text = "File";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(651, 8);
            label3.Name = "label3";
            label3.Size = new Size(109, 30);
            label3.TabIndex = 12;
            label3.Text = "選擇個股";
            // 
            // cboFrameNum
            // 
            cboFrameNum.FormattingEnabled = true;
            cboFrameNum.Items.AddRange(new object[] { "2", "3", "4", "5", "6", "7", "8", "9" });
            cboFrameNum.Location = new Point(67, 5);
            cboFrameNum.Name = "cboFrameNum";
            cboFrameNum.Size = new Size(42, 38);
            cboFrameNum.TabIndex = 11;
            cboFrameNum.Text = "5";
            cboFrameNum.SelectedIndexChanged += cboFrameNum_SelectedIndexChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(415, 9);
            label2.Name = "label2";
            label2.Size = new Size(109, 30);
            label2.TabIndex = 10;
            label2.Text = "股票來源";
            // 
            // btnBack3
            // 
            btnBack3.FlatStyle = FlatStyle.Popup;
            btnBack3.Image = Properties.Resources.BACK3;
            btnBack3.ImageAlign = ContentAlignment.TopLeft;
            btnBack3.Location = new Point(298, 8);
            btnBack3.Name = "btnBack3";
            btnBack3.Size = new Size(30, 25);
            btnBack3.TabIndex = 9;
            btnBack3.UseVisualStyleBackColor = true;
            btnBack3.Click += btnBack3_Click;
            // 
            // btnBack2
            // 
            btnBack2.FlatStyle = FlatStyle.Popup;
            btnBack2.Image = Properties.Resources.BACK2;
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
            btnBack1.Image = Properties.Resources.BACK1;
            btnBack1.ImageAlign = ContentAlignment.TopLeft;
            btnBack1.Location = new Point(244, 8);
            btnBack1.Name = "btnBack1";
            btnBack1.Size = new Size(30, 25);
            btnBack1.TabIndex = 7;
            btnBack1.UseVisualStyleBackColor = true;
            btnBack1.Click += btnBack1_Click;
            // 
            // btnFore1
            // 
            btnFore1.FlatStyle = FlatStyle.Popup;
            btnFore1.Image = Properties.Resources.FORE1;
            btnFore1.ImageAlign = ContentAlignment.TopLeft;
            btnFore1.Location = new Point(218, 8);
            btnFore1.Name = "btnFore1";
            btnFore1.Size = new Size(30, 25);
            btnFore1.TabIndex = 6;
            btnFore1.UseVisualStyleBackColor = true;
            btnFore1.Click += btnFore1_Click;
            // 
            // btnFore2
            // 
            btnFore2.FlatStyle = FlatStyle.Popup;
            btnFore2.Image = Properties.Resources.FORE2;
            btnFore2.ImageAlign = ContentAlignment.TopLeft;
            btnFore2.Location = new Point(192, 8);
            btnFore2.Name = "btnFore2";
            btnFore2.Size = new Size(30, 25);
            btnFore2.TabIndex = 5;
            btnFore2.UseVisualStyleBackColor = true;
            btnFore2.Click += btnFore2_Click;
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
            btnFore3.Image = Properties.Resources.FORE3;
            btnFore3.ImageAlign = ContentAlignment.TopLeft;
            btnFore3.Location = new Point(166, 8);
            btnFore3.Name = "btnFore3";
            btnFore3.Size = new Size(30, 25);
            btnFore3.TabIndex = 3;
            btnFore3.UseVisualStyleBackColor = true;
            btnFore3.Click += btnFore3_Click;
            // 
            // btnZoomOut
            // 
            btnZoomOut.FlatStyle = FlatStyle.Popup;
            btnZoomOut.Image = Properties.Resources.ZOOMOUT;
            btnZoomOut.Location = new Point(140, 8);
            btnZoomOut.Name = "btnZoomOut";
            btnZoomOut.Size = new Size(30, 25);
            btnZoomOut.TabIndex = 2;
            btnZoomOut.UseVisualStyleBackColor = true;
            btnZoomOut.Click += btnZoomOut_Click;
            // 
            // btnZoomIn
            // 
            btnZoomIn.FlatStyle = FlatStyle.Popup;
            btnZoomIn.Image = Properties.Resources.ZOOMIN;
            btnZoomIn.Location = new Point(115, 8);
            btnZoomIn.Name = "btnZoomIn";
            btnZoomIn.Size = new Size(30, 25);
            btnZoomIn.TabIndex = 1;
            btnZoomIn.UseVisualStyleBackColor = true;
            btnZoomIn.Click += btnZoomIn_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(4, 8);
            label1.Name = "label1";
            label1.Size = new Size(85, 30);
            label1.TabIndex = 0;
            label1.Text = "視窗數";
            // 
            // lblMAP1
            // 
            lblMAP1.AutoSize = true;
            lblMAP1.Font = new Font("Arial Narrow", 9.75F);
            lblMAP1.Location = new Point(1031, 80);
            lblMAP1.Name = "lblMAP1";
            lblMAP1.Size = new Size(56, 24);
            lblMAP1.TabIndex = 3;
            lblMAP1.Text = "MAP5";
            // 
            // lblMAP2
            // 
            lblMAP2.AutoSize = true;
            lblMAP2.Font = new Font("Arial Narrow", 9.75F);
            lblMAP2.Location = new Point(1031, 105);
            lblMAP2.Name = "lblMAP2";
            lblMAP2.Size = new Size(65, 24);
            lblMAP2.TabIndex = 4;
            lblMAP2.Text = "MAP10";
            // 
            // lblMAP3
            // 
            lblMAP3.AutoSize = true;
            lblMAP3.Font = new Font("Arial Narrow", 9.75F);
            lblMAP3.Location = new Point(1031, 128);
            lblMAP3.Name = "lblMAP3";
            lblMAP3.Size = new Size(65, 24);
            lblMAP3.TabIndex = 5;
            lblMAP3.Text = "MAP20";
            // 
            // lblMAP4
            // 
            lblMAP4.AutoSize = true;
            lblMAP4.Font = new Font("Arial Narrow", 9.75F);
            lblMAP4.Location = new Point(1031, 152);
            lblMAP4.Name = "lblMAP4";
            lblMAP4.Size = new Size(65, 24);
            lblMAP4.TabIndex = 6;
            lblMAP4.Text = "MAP60";
            // 
            // lblMAP5
            // 
            lblMAP5.AutoSize = true;
            lblMAP5.Font = new Font("Arial Narrow", 9.75F);
            lblMAP5.Location = new Point(1031, 175);
            lblMAP5.Name = "lblMAP5";
            lblMAP5.Size = new Size(74, 24);
            lblMAP5.TabIndex = 7;
            lblMAP5.Text = "MAP120";
            // 
            // lblMAP6
            // 
            lblMAP6.AutoSize = true;
            lblMAP6.Font = new Font("Arial Narrow", 9.75F);
            lblMAP6.Location = new Point(1031, 199);
            lblMAP6.Name = "lblMAP6";
            lblMAP6.Size = new Size(74, 24);
            lblMAP6.TabIndex = 8;
            lblMAP6.Text = "MAP240";
            // 
            // lblMAV1
            // 
            lblMAV1.AutoSize = true;
            lblMAV1.Font = new Font("Arial Narrow", 9.75F);
            lblMAV1.ForeColor = Color.Fuchsia;
            lblMAV1.Location = new Point(1031, 228);
            lblMAV1.Name = "lblMAV1";
            lblMAV1.Size = new Size(56, 24);
            lblMAV1.TabIndex = 9;
            lblMAV1.Text = "MAV5";
            // 
            // lblMAV2
            // 
            lblMAV2.AutoSize = true;
            lblMAV2.Font = new Font("Arial Narrow", 9.75F);
            lblMAV2.ForeColor = Color.Fuchsia;
            lblMAV2.Location = new Point(1031, 248);
            lblMAV2.Name = "lblMAV2";
            lblMAV2.Size = new Size(65, 24);
            lblMAV2.TabIndex = 10;
            lblMAV2.Text = "MAV10";
            // 
            // lblMAV3
            // 
            lblMAV3.AutoSize = true;
            lblMAV3.Font = new Font("Arial Narrow", 9.75F);
            lblMAV3.ForeColor = Color.Fuchsia;
            lblMAV3.Location = new Point(1031, 272);
            lblMAV3.Name = "lblMAV3";
            lblMAV3.Size = new Size(65, 24);
            lblMAV3.TabIndex = 11;
            lblMAV3.Text = "MAV20";
            // 
            // lblMAV4
            // 
            lblMAV4.AutoSize = true;
            lblMAV4.Font = new Font("Arial Narrow", 9.75F);
            lblMAV4.ForeColor = Color.Fuchsia;
            lblMAV4.Location = new Point(1031, 292);
            lblMAV4.Name = "lblMAV4";
            lblMAV4.Size = new Size(65, 24);
            lblMAV4.TabIndex = 12;
            lblMAV4.Text = "MAV60";
            // 
            // lblMAV5
            // 
            lblMAV5.AutoSize = true;
            lblMAV5.Font = new Font("Arial Narrow", 9.75F);
            lblMAV5.ForeColor = Color.Fuchsia;
            lblMAV5.Location = new Point(1031, 312);
            lblMAV5.Name = "lblMAV5";
            lblMAV5.Size = new Size(74, 24);
            lblMAV5.TabIndex = 13;
            lblMAV5.Text = "MAV120";
            // 
            // cntMenuStrip
            // 
            cntMenuStrip.ImageScalingSize = new Size(24, 24);
            cntMenuStrip.Items.AddRange(new ToolStripItem[] { k線ToolStripMenuItem, 量能指標ToolStripMenuItem, 熱門指標ToolStripMenuItem, dFTSToolStripMenuItem, 灰聚類模型ToolStripMenuItem });
            cntMenuStrip.Name = "cntMenuStrip";
            cntMenuStrip.Size = new Size(241, 187);
            // 
            // k線ToolStripMenuItem
            // 
            k線ToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 移動平均線ToolStripMenuItem });
            k線ToolStripMenuItem.Name = "k線ToolStripMenuItem";
            k線ToolStripMenuItem.Size = new Size(240, 30);
            k線ToolStripMenuItem.Text = "K線";
            // 
            // 移動平均線ToolStripMenuItem
            // 
            移動平均線ToolStripMenuItem.Name = "移動平均線ToolStripMenuItem";
            移動平均線ToolStripMenuItem.Size = new Size(200, 34);
            移動平均線ToolStripMenuItem.Text = "移動平均線";
            // 
            // 量能指標ToolStripMenuItem
            // 
            量能指標ToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { VolumeToolStripMenuItem, 融資餘額ToolStripMenuItem, 融資增減ToolStripMenuItem, 融券餘額ToolStripMenuItem, 法人庫存ToolStripMenuItem, 自營商庫存ToolStripMenuItem, 投信庫存ToolStripMenuItem, 外資庫存ToolStripMenuItem });
            量能指標ToolStripMenuItem.Name = "量能指標ToolStripMenuItem";
            量能指標ToolStripMenuItem.Size = new Size(240, 30);
            量能指標ToolStripMenuItem.Text = "量能指標";
            // 
            // VolumeToolStripMenuItem
            // 
            VolumeToolStripMenuItem.Name = "VolumeToolStripMenuItem";
            VolumeToolStripMenuItem.Size = new Size(200, 34);
            VolumeToolStripMenuItem.Text = "成交量";
            VolumeToolStripMenuItem.Click += VolumeToolStripMenuItem_Click;
            // 
            // 融資餘額ToolStripMenuItem
            // 
            融資餘額ToolStripMenuItem.Name = "融資餘額ToolStripMenuItem";
            融資餘額ToolStripMenuItem.Size = new Size(200, 34);
            融資餘額ToolStripMenuItem.Text = "融資餘額";
            // 
            // 融資增減ToolStripMenuItem
            // 
            融資增減ToolStripMenuItem.Name = "融資增減ToolStripMenuItem";
            融資增減ToolStripMenuItem.Size = new Size(200, 34);
            融資增減ToolStripMenuItem.Text = "融資增減";
            // 
            // 融券餘額ToolStripMenuItem
            // 
            融券餘額ToolStripMenuItem.Name = "融券餘額ToolStripMenuItem";
            融券餘額ToolStripMenuItem.Size = new Size(200, 34);
            融券餘額ToolStripMenuItem.Text = "融券餘額";
            // 
            // 法人庫存ToolStripMenuItem
            // 
            法人庫存ToolStripMenuItem.Name = "法人庫存ToolStripMenuItem";
            法人庫存ToolStripMenuItem.Size = new Size(200, 34);
            法人庫存ToolStripMenuItem.Text = "法人庫存";
            // 
            // 自營商庫存ToolStripMenuItem
            // 
            自營商庫存ToolStripMenuItem.Name = "自營商庫存ToolStripMenuItem";
            自營商庫存ToolStripMenuItem.Size = new Size(200, 34);
            自營商庫存ToolStripMenuItem.Text = "自營商庫存";
            // 
            // 投信庫存ToolStripMenuItem
            // 
            投信庫存ToolStripMenuItem.Name = "投信庫存ToolStripMenuItem";
            投信庫存ToolStripMenuItem.Size = new Size(200, 34);
            投信庫存ToolStripMenuItem.Text = "投信庫存";
            // 
            // 外資庫存ToolStripMenuItem
            // 
            外資庫存ToolStripMenuItem.Name = "外資庫存ToolStripMenuItem";
            外資庫存ToolStripMenuItem.Size = new Size(200, 34);
            外資庫存ToolStripMenuItem.Text = "外資庫存";
            // 
            // 熱門指標ToolStripMenuItem
            // 
            熱門指標ToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { KDToolStripMenuItem, MACDToolStripMenuItem, RSIToolStripMenuItem, stochRSIToolStripMenuItem, wildersRSIToolStripMenuItem, WMSToolStripMenuItem, WMSCToolStripMenuItem, EMAToolStripMenuItem, BIASToolStripMenuItem });
            熱門指標ToolStripMenuItem.Name = "熱門指標ToolStripMenuItem";
            熱門指標ToolStripMenuItem.Size = new Size(240, 30);
            熱門指標ToolStripMenuItem.Text = "熱門指標";
            // 
            // KDToolStripMenuItem
            // 
            KDToolStripMenuItem.Name = "KDToolStripMenuItem";
            KDToolStripMenuItem.Size = new Size(270, 34);
            KDToolStripMenuItem.Text = "KD";
            // 
            // MACDToolStripMenuItem
            // 
            MACDToolStripMenuItem.Name = "MACDToolStripMenuItem";
            MACDToolStripMenuItem.Size = new Size(270, 34);
            MACDToolStripMenuItem.Text = "MACD";
            // 
            // RSIToolStripMenuItem
            // 
            RSIToolStripMenuItem.Name = "RSIToolStripMenuItem";
            RSIToolStripMenuItem.Size = new Size(270, 34);
            RSIToolStripMenuItem.Text = "RSI";
            // 
            // stochRSIToolStripMenuItem
            // 
            stochRSIToolStripMenuItem.Name = "stochRSIToolStripMenuItem";
            stochRSIToolStripMenuItem.Size = new Size(270, 34);
            stochRSIToolStripMenuItem.Text = "Stoch RSI";
            // 
            // wildersRSIToolStripMenuItem
            // 
            wildersRSIToolStripMenuItem.Name = "wildersRSIToolStripMenuItem";
            wildersRSIToolStripMenuItem.Size = new Size(270, 34);
            wildersRSIToolStripMenuItem.Text = "Wilder's RSI";
            // 
            // WMSToolStripMenuItem
            // 
            WMSToolStripMenuItem.Name = "WMSToolStripMenuItem";
            WMSToolStripMenuItem.Size = new Size(270, 34);
            WMSToolStripMenuItem.Text = "WMS";
            WMSToolStripMenuItem.Click += WMSToolStripMenuItem_Click;
            // 
            // WMSCToolStripMenuItem
            // 
            WMSCToolStripMenuItem.Name = "WMSCToolStripMenuItem";
            WMSCToolStripMenuItem.Size = new Size(270, 34);
            WMSCToolStripMenuItem.Text = "WMSC";
            // 
            // EMAToolStripMenuItem
            // 
            EMAToolStripMenuItem.Name = "EMAToolStripMenuItem";
            EMAToolStripMenuItem.Size = new Size(270, 34);
            EMAToolStripMenuItem.Text = "EMA";
            // 
            // BIASToolStripMenuItem
            // 
            BIASToolStripMenuItem.Name = "BIASToolStripMenuItem";
            BIASToolStripMenuItem.Size = new Size(270, 34);
            BIASToolStripMenuItem.Text = "BIAS";
            BIASToolStripMenuItem.Click += BIASToolStripMenuItem_Click;
            // 
            // dFTSToolStripMenuItem
            // 
            dFTSToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { trendToolStripMenuItem, qMToolStripMenuItem, wLSTToolStripMenuItem, wCYToolStripMenuItem, wLWToolStripMenuItem, dIFFToolStripMenuItem, dCToolStripMenuItem, holdToolStripMenuItem, profitToolStripMenuItem, cYToolStripMenuItem, signalsToolStripMenuItem });
            dFTSToolStripMenuItem.Name = "dFTSToolStripMenuItem";
            dFTSToolStripMenuItem.Size = new Size(240, 30);
            dFTSToolStripMenuItem.Text = "DFTS";
            // 
            // trendToolStripMenuItem
            // 
            trendToolStripMenuItem.Name = "trendToolStripMenuItem";
            trendToolStripMenuItem.Size = new Size(170, 34);
            trendToolStripMenuItem.Text = "Trend";
            // 
            // qMToolStripMenuItem
            // 
            qMToolStripMenuItem.Name = "qMToolStripMenuItem";
            qMToolStripMenuItem.Size = new Size(170, 34);
            qMToolStripMenuItem.Text = "QM";
            // 
            // wLSTToolStripMenuItem
            // 
            wLSTToolStripMenuItem.Name = "wLSTToolStripMenuItem";
            wLSTToolStripMenuItem.Size = new Size(170, 34);
            wLSTToolStripMenuItem.Text = "WLST";
            // 
            // wCYToolStripMenuItem
            // 
            wCYToolStripMenuItem.Name = "wCYToolStripMenuItem";
            wCYToolStripMenuItem.Size = new Size(170, 34);
            wCYToolStripMenuItem.Text = "WCY";
            // 
            // wLWToolStripMenuItem
            // 
            wLWToolStripMenuItem.Name = "wLWToolStripMenuItem";
            wLWToolStripMenuItem.Size = new Size(170, 34);
            wLWToolStripMenuItem.Text = "WLW";
            // 
            // dIFFToolStripMenuItem
            // 
            dIFFToolStripMenuItem.Name = "dIFFToolStripMenuItem";
            dIFFToolStripMenuItem.Size = new Size(170, 34);
            dIFFToolStripMenuItem.Text = "DIFF";
            // 
            // dCToolStripMenuItem
            // 
            dCToolStripMenuItem.Name = "dCToolStripMenuItem";
            dCToolStripMenuItem.Size = new Size(170, 34);
            dCToolStripMenuItem.Text = "DC";
            // 
            // holdToolStripMenuItem
            // 
            holdToolStripMenuItem.Name = "holdToolStripMenuItem";
            holdToolStripMenuItem.Size = new Size(170, 34);
            holdToolStripMenuItem.Text = "CY";
            // 
            // profitToolStripMenuItem
            // 
            profitToolStripMenuItem.Name = "profitToolStripMenuItem";
            profitToolStripMenuItem.Size = new Size(170, 34);
            profitToolStripMenuItem.Text = "Hold";
            // 
            // cYToolStripMenuItem
            // 
            cYToolStripMenuItem.Name = "cYToolStripMenuItem";
            cYToolStripMenuItem.Size = new Size(170, 34);
            cYToolStripMenuItem.Text = "Profit";
            // 
            // signalsToolStripMenuItem
            // 
            signalsToolStripMenuItem.Name = "signalsToolStripMenuItem";
            signalsToolStripMenuItem.Size = new Size(170, 34);
            signalsToolStripMenuItem.Text = "Signals";
            // 
            // 灰聚類模型ToolStripMenuItem
            // 
            灰聚類模型ToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 切割區間ToolStripMenuItem, 聚類結果ToolStripMenuItem });
            灰聚類模型ToolStripMenuItem.Name = "灰聚類模型ToolStripMenuItem";
            灰聚類模型ToolStripMenuItem.Size = new Size(240, 30);
            灰聚類模型ToolStripMenuItem.Text = "灰聚類模型";
            // 
            // 切割區間ToolStripMenuItem
            // 
            切割區間ToolStripMenuItem.Name = "切割區間ToolStripMenuItem";
            切割區間ToolStripMenuItem.Size = new Size(182, 34);
            切割區間ToolStripMenuItem.Text = "切割區間";
            // 
            // 聚類結果ToolStripMenuItem
            // 
            聚類結果ToolStripMenuItem.Name = "聚類結果ToolStripMenuItem";
            聚類結果ToolStripMenuItem.Size = new Size(182, 34);
            聚類結果ToolStripMenuItem.Text = "聚類結果";
            // 
            // frmStocksPGM
            // 
            AutoScaleDimensions = new SizeF(14F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1184, 761);
            Controls.Add(lblMAV5);
            Controls.Add(lblMAV4);
            Controls.Add(lblMAV3);
            Controls.Add(lblMAV2);
            Controls.Add(lblMAV1);
            Controls.Add(lblMAP6);
            Controls.Add(lblMAP5);
            Controls.Add(lblMAP4);
            Controls.Add(lblMAP3);
            Controls.Add(lblMAP2);
            Controls.Add(lblMAP1);
            Controls.Add(pnlStocksBar);
            Controls.Add(mnuStocksList);
            Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 136);
            MainMenuStrip = mnuStocksList;
            Margin = new Padding(4);
            Name = "frmStocksPGM";
            Text = "Stocks Test Form";
            Load += frmStocksPGM_Load;
            Paint += frmStocksPGM_Paint;
            MouseMove += frmStocksPGM_MouseMove;
            Resize += frmStocksPGM_Resize;
            mnuStocksList.ResumeLayout(false);
            mnuStocksList.PerformLayout();
            pnlStocksBar.ResumeLayout(false);
            pnlStocksBar.PerformLayout();
            cntMenuStrip.ResumeLayout(false);
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
        private Label lblMAP1;
        private Label lblMAP2;
        private Label lblMAP3;
        private Label lblMAP4;
        private Label lblMAP5;
        private Label lblMAP6;
        private Label lblMAV1;
        private Label lblMAV2;
        private Label lblMAV3;
        private Label lblMAV4;
        private Label lblMAV5;
        private ContextMenuStrip cntMenuStrip;
        private ToolStripMenuItem k線ToolStripMenuItem;
        private ToolStripMenuItem 量能指標ToolStripMenuItem;
        private ToolStripMenuItem VolumeToolStripMenuItem;
        private ToolStripMenuItem 融資餘額ToolStripMenuItem;
        private ToolStripMenuItem 融資增減ToolStripMenuItem;
        private ToolStripMenuItem 融券餘額ToolStripMenuItem;
        private ToolStripMenuItem 法人庫存ToolStripMenuItem;
        private ToolStripMenuItem 熱門指標ToolStripMenuItem;
        private ToolStripMenuItem dFTSToolStripMenuItem;
        private ToolStripMenuItem 灰聚類模型ToolStripMenuItem;
        private ToolStripMenuItem 移動平均線ToolStripMenuItem;
        private ToolStripMenuItem 自營商庫存ToolStripMenuItem;
        private ToolStripMenuItem 投信庫存ToolStripMenuItem;
        private ToolStripMenuItem 外資庫存ToolStripMenuItem;
        private ToolStripMenuItem KDToolStripMenuItem;
        private ToolStripMenuItem MACDToolStripMenuItem;
        private ToolStripMenuItem RSIToolStripMenuItem;
        private ToolStripMenuItem stochRSIToolStripMenuItem;
        private ToolStripMenuItem wildersRSIToolStripMenuItem;
        private ToolStripMenuItem WMSToolStripMenuItem;
        private ToolStripMenuItem WMSCToolStripMenuItem;
        private ToolStripMenuItem EMAToolStripMenuItem;
        private ToolStripMenuItem BIASToolStripMenuItem;
        private ToolStripMenuItem trendToolStripMenuItem;
        private ToolStripMenuItem qMToolStripMenuItem;
        private ToolStripMenuItem wLSTToolStripMenuItem;
        private ToolStripMenuItem wCYToolStripMenuItem;
        private ToolStripMenuItem wLWToolStripMenuItem;
        private ToolStripMenuItem dIFFToolStripMenuItem;
        private ToolStripMenuItem dCToolStripMenuItem;
        private ToolStripMenuItem holdToolStripMenuItem;
        private ToolStripMenuItem profitToolStripMenuItem;
        private ToolStripMenuItem cYToolStripMenuItem;
        private ToolStripMenuItem signalsToolStripMenuItem;
        private ToolStripMenuItem 切割區間ToolStripMenuItem;
        private ToolStripMenuItem 聚類結果ToolStripMenuItem;
    }
}
