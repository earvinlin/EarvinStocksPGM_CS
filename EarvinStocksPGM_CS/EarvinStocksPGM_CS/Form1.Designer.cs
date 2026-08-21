namespace EarvinStocksPGM_CS
{
    partial class frmEarvinStocks
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
            btnTest = new Button();
            lblTest = new Label();
            SuspendLayout();
            // 
            // btnTest
            // 
            btnTest.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 136);
            btnTest.Location = new Point(809, 76);
            btnTest.Margin = new Padding(4, 4, 4, 4);
            btnTest.Name = "btnTest";
            btnTest.Size = new Size(136, 40);
            btnTest.TabIndex = 0;
            btnTest.Text = "移動我!";
            btnTest.UseVisualStyleBackColor = true;
            btnTest.Click += btnTest_Click;
            // 
            // lblTest
            // 
            lblTest.AutoSize = true;
            lblTest.Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 136);
            lblTest.Location = new Point(171, 88);
            lblTest.Margin = new Padding(4, 0, 4, 0);
            lblTest.Name = "lblTest";
            lblTest.Size = new Size(88, 20);
            lblTest.TabIndex = 1;
            lblTest.Text = "測試文字!!!";
            // 
            // frmEarvinStocks
            // 
            AutoScaleDimensions = new SizeF(10F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1143, 600);
            Controls.Add(lblTest);
            Controls.Add(btnTest);
            Font = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 136);
            Margin = new Padding(4, 4, 4, 4);
            Name = "frmEarvinStocks";
            Text = "EarvinStocks";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnTest;
        private Label lblTest;
    }
}
