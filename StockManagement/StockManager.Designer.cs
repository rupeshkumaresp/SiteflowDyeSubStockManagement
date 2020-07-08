namespace nsDyeSubStockManagement
{
    partial class StockManager
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnDownload = new System.Windows.Forms.Button();
            this.tabDownload = new System.Windows.Forms.TabControl();
            this.tabPageDownload = new System.Windows.Forms.TabPage();
            this.lblStatusDownload = new System.Windows.Forms.Label();
            this.tabPageUpload = new System.Windows.Forms.TabPage();
            this.groupBoxUpdateStock = new System.Windows.Forms.GroupBox();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnUpload = new System.Windows.Forms.Button();
            this.txtBoxOutput = new System.Windows.Forms.TextBox();
            this.labelUploadMessage = new System.Windows.Forms.Label();
            this.tabDownload.SuspendLayout();
            this.tabPageDownload.SuspendLayout();
            this.tabPageUpload.SuspendLayout();
            this.groupBoxUpdateStock.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnDownload
            // 
            this.btnDownload.Location = new System.Drawing.Point(95, 67);
            this.btnDownload.Name = "btnDownload";
            this.btnDownload.Size = new System.Drawing.Size(198, 37);
            this.btnDownload.TabIndex = 0;
            this.btnDownload.Text = "Download Stock(.xlsx)";
            this.btnDownload.UseVisualStyleBackColor = true;
            this.btnDownload.Click += new System.EventHandler(this.btnDownload_Click);
            // 
            // tabDownload
            // 
            this.tabDownload.Controls.Add(this.tabPageDownload);
            this.tabDownload.Controls.Add(this.tabPageUpload);
            this.tabDownload.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabDownload.Location = new System.Drawing.Point(27, 24);
            this.tabDownload.Name = "tabDownload";
            this.tabDownload.SelectedIndex = 0;
            this.tabDownload.Size = new System.Drawing.Size(404, 213);
            this.tabDownload.TabIndex = 2;
            this.tabDownload.SelectedIndexChanged += new System.EventHandler(this.tabDownload_SelectedIndexChanged);
            // 
            // tabPageDownload
            // 
            this.tabPageDownload.Controls.Add(this.lblStatusDownload);
            this.tabPageDownload.Controls.Add(this.btnDownload);
            this.tabPageDownload.Location = new System.Drawing.Point(4, 24);
            this.tabPageDownload.Name = "tabPageDownload";
            this.tabPageDownload.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageDownload.Size = new System.Drawing.Size(396, 185);
            this.tabPageDownload.TabIndex = 0;
            this.tabPageDownload.Text = "View Stocks";
            this.tabPageDownload.UseVisualStyleBackColor = true;
            // 
            // lblStatusDownload
            // 
            this.lblStatusDownload.AutoSize = true;
            this.lblStatusDownload.Location = new System.Drawing.Point(92, 137);
            this.lblStatusDownload.Name = "lblStatusDownload";
            this.lblStatusDownload.Size = new System.Drawing.Size(0, 15);
            this.lblStatusDownload.TabIndex = 1;
            // 
            // tabPageUpload
            // 
            this.tabPageUpload.Controls.Add(this.groupBoxUpdateStock);
            this.tabPageUpload.Location = new System.Drawing.Point(4, 24);
            this.tabPageUpload.Name = "tabPageUpload";
            this.tabPageUpload.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageUpload.Size = new System.Drawing.Size(396, 185);
            this.tabPageUpload.TabIndex = 1;
            this.tabPageUpload.Text = "Update Stocks";
            this.tabPageUpload.UseVisualStyleBackColor = true;
            // 
            // groupBoxUpdateStock
            // 
            this.groupBoxUpdateStock.Controls.Add(this.labelUploadMessage);
            this.groupBoxUpdateStock.Controls.Add(this.btnClear);
            this.groupBoxUpdateStock.Controls.Add(this.btnBrowse);
            this.groupBoxUpdateStock.Controls.Add(this.lblStatus);
            this.groupBoxUpdateStock.Controls.Add(this.btnUpload);
            this.groupBoxUpdateStock.Controls.Add(this.txtBoxOutput);
            this.groupBoxUpdateStock.Location = new System.Drawing.Point(6, 6);
            this.groupBoxUpdateStock.Name = "groupBoxUpdateStock";
            this.groupBoxUpdateStock.Size = new System.Drawing.Size(384, 175);
            this.groupBoxUpdateStock.TabIndex = 17;
            this.groupBoxUpdateStock.TabStop = false;
            this.groupBoxUpdateStock.Text = "Browse Stock Spreadsheet and Upload";
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(80, 80);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(80, 28);
            this.btnClear.TabIndex = 17;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(307, 40);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(71, 23);
            this.btnBrowse.TabIndex = 15;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(66, 146);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 15);
            this.lblStatus.TabIndex = 16;
            // 
            // btnUpload
            // 
            this.btnUpload.Location = new System.Drawing.Point(178, 79);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(123, 29);
            this.btnUpload.TabIndex = 13;
            this.btnUpload.Text = "Upload Stock";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // txtBoxOutput
            // 
            this.txtBoxOutput.Enabled = false;
            this.txtBoxOutput.Location = new System.Drawing.Point(6, 42);
            this.txtBoxOutput.Name = "txtBoxOutput";
            this.txtBoxOutput.Size = new System.Drawing.Size(295, 21);
            this.txtBoxOutput.TabIndex = 14;
            // 
            // labelUploadMessage
            // 
            this.labelUploadMessage.AutoSize = true;
            this.labelUploadMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelUploadMessage.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.labelUploadMessage.Location = new System.Drawing.Point(6, 115);
            this.labelUploadMessage.Name = "labelUploadMessage";
            this.labelUploadMessage.Size = new System.Drawing.Size(379, 15);
            this.labelUploadMessage.TabIndex = 18;
            this.labelUploadMessage.Text = "(Upload will update all the stocks data including week consumtption)";
            // 
            // StockManager
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(464, 291);
            this.Controls.Add(this.tabDownload);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(480, 330);
            this.MinimumSize = new System.Drawing.Size(480, 330);
            this.Name = "StockManager";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Site flow Dye Sub Stock Management";
            this.tabDownload.ResumeLayout(false);
            this.tabPageDownload.ResumeLayout(false);
            this.tabPageDownload.PerformLayout();
            this.tabPageUpload.ResumeLayout(false);
            this.groupBoxUpdateStock.ResumeLayout(false);
            this.groupBoxUpdateStock.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnDownload;
        private System.Windows.Forms.TabControl tabDownload;
        private System.Windows.Forms.TabPage tabPageDownload;
        private System.Windows.Forms.TabPage tabPageUpload;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.TextBox txtBoxOutput;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.GroupBox groupBoxUpdateStock;
        private System.Windows.Forms.Label lblStatusDownload;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Label labelUploadMessage;
    }
}

