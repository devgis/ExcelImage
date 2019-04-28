namespace LYF.ExcelImage
{
    partial class Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btNewFile2003 = new System.Windows.Forms.Button();
            this.btNewFile2007 = new System.Windows.Forms.Button();
            this.btNewWord2007 = new System.Windows.Forms.Button();
            this.btNewWord2003 = new System.Windows.Forms.Button();
            this.btExcelImage = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lbInfo = new System.Windows.Forms.Label();
            this.btExcelOffice = new System.Windows.Forms.Button();
            this.btStop = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btNewFile2003
            // 
            this.btNewFile2003.Location = new System.Drawing.Point(44, 29);
            this.btNewFile2003.Name = "btNewFile2003";
            this.btNewFile2003.Size = new System.Drawing.Size(118, 23);
            this.btNewFile2003.TabIndex = 0;
            this.btNewFile2003.Text = "创建Excel2003文件";
            this.btNewFile2003.UseVisualStyleBackColor = true;
            this.btNewFile2003.Visible = false;
            // 
            // btNewFile2007
            // 
            this.btNewFile2007.Location = new System.Drawing.Point(44, 69);
            this.btNewFile2007.Name = "btNewFile2007";
            this.btNewFile2007.Size = new System.Drawing.Size(118, 23);
            this.btNewFile2007.TabIndex = 1;
            this.btNewFile2007.Text = "创建Excel2007文件";
            this.btNewFile2007.UseVisualStyleBackColor = true;
            this.btNewFile2007.Visible = false;
            // 
            // btNewWord2007
            // 
            this.btNewWord2007.Location = new System.Drawing.Point(210, 69);
            this.btNewWord2007.Name = "btNewWord2007";
            this.btNewWord2007.Size = new System.Drawing.Size(118, 23);
            this.btNewWord2007.TabIndex = 3;
            this.btNewWord2007.Text = "创建Word2007文件";
            this.btNewWord2007.UseVisualStyleBackColor = true;
            this.btNewWord2007.Visible = false;
            // 
            // btNewWord2003
            // 
            this.btNewWord2003.Location = new System.Drawing.Point(210, 29);
            this.btNewWord2003.Name = "btNewWord2003";
            this.btNewWord2003.Size = new System.Drawing.Size(118, 23);
            this.btNewWord2003.TabIndex = 2;
            this.btNewWord2003.Text = "创建Word2003文件";
            this.btNewWord2003.UseVisualStyleBackColor = true;
            this.btNewWord2003.Visible = false;
            // 
            // btExcelImage
            // 
            this.btExcelImage.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btExcelImage.Location = new System.Drawing.Point(105, 178);
            this.btExcelImage.Name = "btExcelImage";
            this.btExcelImage.Size = new System.Drawing.Size(124, 34);
            this.btExcelImage.TabIndex = 4;
            this.btExcelImage.Text = "创建Excel图片";
            this.btExcelImage.UseVisualStyleBackColor = true;
            this.btExcelImage.Visible = false;
            this.btExcelImage.Click += new System.EventHandler(this.btExcelImage_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(65, 98);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(337, 23);
            this.progressBar.TabIndex = 5;
            // 
            // lbInfo
            // 
            this.lbInfo.AutoSize = true;
            this.lbInfo.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lbInfo.Location = new System.Drawing.Point(65, 128);
            this.lbInfo.Name = "lbInfo";
            this.lbInfo.Size = new System.Drawing.Size(0, 21);
            this.lbInfo.TabIndex = 6;
            // 
            // btExcelOffice
            // 
            this.btExcelOffice.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btExcelOffice.Location = new System.Drawing.Point(105, 178);
            this.btExcelOffice.Name = "btExcelOffice";
            this.btExcelOffice.Size = new System.Drawing.Size(124, 34);
            this.btExcelOffice.TabIndex = 7;
            this.btExcelOffice.Text = "创建Excel图片";
            this.btExcelOffice.UseVisualStyleBackColor = true;
            this.btExcelOffice.Click += new System.EventHandler(this.btExcelOffice_Click);
            // 
            // btStop
            // 
            this.btStop.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btStop.Location = new System.Drawing.Point(235, 178);
            this.btStop.Name = "btStop";
            this.btStop.Size = new System.Drawing.Size(124, 34);
            this.btStop.TabIndex = 8;
            this.btStop.Text = "停止";
            this.btStop.UseVisualStyleBackColor = true;
            this.btStop.Click += new System.EventHandler(this.btStop_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(466, 279);
            this.Controls.Add(this.btStop);
            this.Controls.Add(this.btExcelOffice);
            this.Controls.Add(this.lbInfo);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btExcelImage);
            this.Controls.Add(this.btNewWord2007);
            this.Controls.Add(this.btNewWord2003);
            this.Controls.Add(this.btNewFile2007);
            this.Controls.Add(this.btNewFile2003);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "Main";
            this.Text = "Excel画";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Main_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btNewFile2003;
        private System.Windows.Forms.Button btNewFile2007;
        private System.Windows.Forms.Button btNewWord2007;
        private System.Windows.Forms.Button btNewWord2003;
        private System.Windows.Forms.Button btExcelImage;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lbInfo;
        private System.Windows.Forms.Button btExcelOffice;
        private System.Windows.Forms.Button btStop;
    }
}

