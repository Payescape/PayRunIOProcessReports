namespace PayRunIOProcessReports
{
    partial class Form1
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
            this.btnProduceReports = new DevExpress.XtraEditors.SimpleButton();
            this.SuspendLayout();
            // 
            // btnProduceReports
            // 
            this.btnProduceReports.Location = new System.Drawing.Point(252, 122);
            this.btnProduceReports.Name = "btnProduceReports";
            this.btnProduceReports.Size = new System.Drawing.Size(107, 23);
            this.btnProduceReports.TabIndex = 0;
            this.btnProduceReports.Text = "Produce Reports";
            this.btnProduceReports.Click += new System.EventHandler(this.btnProduceReports_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 278);
            this.Controls.Add(this.btnProduceReports);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.SimpleButton btnProduceReports;
    }
}

