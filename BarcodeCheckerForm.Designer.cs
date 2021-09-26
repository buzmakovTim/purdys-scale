
namespace PortMainScaleTest
{
    partial class BarcodeCheckerForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BarcodeCheckerForm));
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxBarCodeChecker = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft YaHei", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(37, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(327, 35);
            this.label1.TabIndex = 0;
            this.label1.Text = "Please scan the barcode";
            // 
            // textBoxBarCodeChecker
            // 
            this.textBoxBarCodeChecker.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxBarCodeChecker.Location = new System.Drawing.Point(116, 97);
            this.textBoxBarCodeChecker.Name = "textBoxBarCodeChecker";
            this.textBoxBarCodeChecker.Size = new System.Drawing.Size(173, 38);
            this.textBoxBarCodeChecker.TabIndex = 1;
            this.textBoxBarCodeChecker.KeyUp += new System.Windows.Forms.KeyEventHandler(this.textBoxBarCodeChecker_KeyUp);
            // 
            // BarcodeCheckerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(398, 177);
            this.Controls.Add(this.textBoxBarCodeChecker);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "BarcodeCheckerForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Barcode checker";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.BarCodeChecker_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxBarCodeChecker;
    }
}