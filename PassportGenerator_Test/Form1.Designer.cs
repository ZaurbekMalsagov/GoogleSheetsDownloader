namespace GoogleSheetsDownloader {
    partial class Form1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.button1 = new System.Windows.Forms.Button();
            this.txtBxStartRange = new System.Windows.Forms.TextBox();
            this.txtBxEndRange = new System.Windows.Forms.TextBox();
            this.txtBxListName = new System.Windows.Forms.TextBox();
            this.txtBxExcelFile = new System.Windows.Forms.TextBox();
            this.btnReset = new System.Windows.Forms.Button();
            this.txtBxIdTable = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(95, 141);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(156, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Import data from Google";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtBxStartRange
            // 
            this.txtBxStartRange.Location = new System.Drawing.Point(201, 115);
            this.txtBxStartRange.Name = "txtBxStartRange";
            this.txtBxStartRange.Size = new System.Drawing.Size(100, 20);
            this.txtBxStartRange.TabIndex = 1;
            this.txtBxStartRange.Text = "A3";
            // 
            // txtBxEndRange
            // 
            this.txtBxEndRange.Location = new System.Drawing.Point(307, 115);
            this.txtBxEndRange.Name = "txtBxEndRange";
            this.txtBxEndRange.Size = new System.Drawing.Size(100, 20);
            this.txtBxEndRange.TabIndex = 2;
            this.txtBxEndRange.Text = "J3";
            // 
            // txtBxListName
            // 
            this.txtBxListName.Location = new System.Drawing.Point(95, 115);
            this.txtBxListName.Name = "txtBxListName";
            this.txtBxListName.Size = new System.Drawing.Size(100, 20);
            this.txtBxListName.TabIndex = 3;
            // 
            // txtBxExcelFile
            // 
            this.txtBxExcelFile.Location = new System.Drawing.Point(95, 170);
            this.txtBxExcelFile.Name = "txtBxExcelFile";
            this.txtBxExcelFile.Size = new System.Drawing.Size(312, 20);
            this.txtBxExcelFile.TabIndex = 4;
            this.txtBxExcelFile.Text = "C:\\WORK\\PROJECTS\\GoogleSheetsDownloader\\Устройства.xlsx";
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(257, 141);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(150, 23);
            this.btnReset.TabIndex = 5;
            this.btnReset.Text = "Reset Excel File";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // txtBxIdTable
            // 
            this.txtBxIdTable.Location = new System.Drawing.Point(95, 89);
            this.txtBxIdTable.Name = "txtBxIdTable";
            this.txtBxIdTable.Size = new System.Drawing.Size(312, 20);
            this.txtBxIdTable.TabIndex = 6;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(550, 323);
            this.Controls.Add(this.txtBxIdTable);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.txtBxExcelFile);
            this.Controls.Add(this.txtBxListName);
            this.Controls.Add(this.txtBxEndRange);
            this.Controls.Add(this.txtBxStartRange);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtBxEndRange;
        internal System.Windows.Forms.TextBox txtBxStartRange;
        private System.Windows.Forms.TextBox txtBxListName;
        private System.Windows.Forms.TextBox txtBxExcelFile;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.TextBox txtBxIdTable;
    }
}

