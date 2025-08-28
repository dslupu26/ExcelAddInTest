namespace ExcelAddInTest
{
    partial class DebugPane
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
            this.SuspendLayout();
            // 
            // CluOutputBox
            // 
            this.CluOutputBox.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CluOutputBox.Location = new System.Drawing.Point(23, 28);
            this.CluOutputBox.Multiline = true;
            this.CluOutputBox.Name = "CluOutputBox";
            this.CluOutputBox.ReadOnly = true;
            this.CluOutputBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.CluOutputBox.Size = new System.Drawing.Size(500, 506);
            this.CluOutputBox.TabIndex = 0;
            // 
            // DebugPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.CluOutputBox);
            this.Name = "DebugPane";
            this.Size = new System.Drawing.Size(545, 555);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox CluOutputBox;
    }
}
