namespace DataExport
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
            this.btExrpot = new System.Windows.Forms.Button();
            this.tbExportPath = new System.Windows.Forms.TextBox();
            this.tbImportPath = new System.Windows.Forms.TextBox();
            this.btImport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btExrpot
            // 
            this.btExrpot.Location = new System.Drawing.Point(373, 76);
            this.btExrpot.Name = "btExrpot";
            this.btExrpot.Size = new System.Drawing.Size(75, 23);
            this.btExrpot.TabIndex = 0;
            this.btExrpot.Text = "导出数据";
            this.btExrpot.UseVisualStyleBackColor = true;
            this.btExrpot.Click += new System.EventHandler(this.btExrpot_Click);
            // 
            // tbExportPath
            // 
            this.tbExportPath.Location = new System.Drawing.Point(63, 76);
            this.tbExportPath.Name = "tbExportPath";
            this.tbExportPath.Size = new System.Drawing.Size(274, 21);
            this.tbExportPath.TabIndex = 1;
            // 
            // tbImportPath
            // 
            this.tbImportPath.Location = new System.Drawing.Point(63, 133);
            this.tbImportPath.Name = "tbImportPath";
            this.tbImportPath.Size = new System.Drawing.Size(274, 21);
            this.tbImportPath.TabIndex = 3;
            // 
            // btImport
            // 
            this.btImport.Location = new System.Drawing.Point(373, 133);
            this.btImport.Name = "btImport";
            this.btImport.Size = new System.Drawing.Size(75, 23);
            this.btImport.TabIndex = 2;
            this.btImport.Text = "导入数据";
            this.btImport.UseVisualStyleBackColor = true;
            this.btImport.Click += new System.EventHandler(this.btImport_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(517, 261);
            this.Controls.Add(this.tbImportPath);
            this.Controls.Add(this.btImport);
            this.Controls.Add(this.tbExportPath);
            this.Controls.Add(this.btExrpot);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "Main";
            this.Text = "数据导出";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btExrpot;
        private System.Windows.Forms.TextBox tbExportPath;
        private System.Windows.Forms.TextBox tbImportPath;
        private System.Windows.Forms.Button btImport;
    }
}

