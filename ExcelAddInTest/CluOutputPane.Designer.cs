namespace ExcelAddInTest
{
    partial class CluOutputPane
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.CluOutputBox = new System.Windows.Forms.TextBox();
            this.CluOutputLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // CluOutputBox
            // 
            this.CluOutputBox.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CluOutputBox.Location = new System.Drawing.Point(23, 61);
            this.CluOutputBox.Multiline = true;
            this.CluOutputBox.Name = "CluOutputBox";
            this.CluOutputBox.ReadOnly = true;
            this.CluOutputBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.CluOutputBox.Size = new System.Drawing.Size(482, 473);
            this.CluOutputBox.TabIndex = 0;
            // 
            // CluOutputLabel
            // 
            this.CluOutputLabel.AutoSize = true;
            this.CluOutputLabel.Font = new System.Drawing.Font("Microsoft YaHei", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CluOutputLabel.Location = new System.Drawing.Point(17, 15);
            this.CluOutputLabel.Name = "CluOutputLabel";
            this.CluOutputLabel.Size = new System.Drawing.Size(135, 31);
            this.CluOutputLabel.TabIndex = 1;
            this.CluOutputLabel.Text = "CluOutput";
            // 
            // CluOutputPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.CluOutputLabel);
            this.Controls.Add(this.CluOutputBox);
            this.Name = "CluOutputPane";
            this.Size = new System.Drawing.Size(545, 555);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox CluOutputBox;
        private System.Windows.Forms.Label CluOutputLabel;
    }
}
