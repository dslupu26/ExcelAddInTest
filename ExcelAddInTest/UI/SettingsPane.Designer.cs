using System.Drawing;
using System.Windows.Forms;

namespace ExcelAddInTest
{
    partial class SettingsPane
    {
        private System.ComponentModel.IContainer components = null;
        private CheckBox chkContinuous;
        private TrackBar tbSeconds;
        private Label lblSeconds;
        private CheckBox chkSingle;


        // NEW: numeric guides
        private TableLayoutPanel tlpTicks;
        private Label lbl1;
        private Label lbl5;
        private Label lbl10;
        private Label lbl15;
        private Label lbl20;
        private Label lbl25;
        private Label lbl30;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.chkContinuous = new System.Windows.Forms.CheckBox();
            this.tbSeconds = new System.Windows.Forms.TrackBar();
            this.lblSeconds = new System.Windows.Forms.Label();
            this.chkSingle = new System.Windows.Forms.CheckBox();

            // NEW: create table + labels
            this.tlpTicks = new System.Windows.Forms.TableLayoutPanel();
            this.lbl1 = new System.Windows.Forms.Label();
            this.lbl5 = new System.Windows.Forms.Label();
            this.lbl10 = new System.Windows.Forms.Label();
            this.lbl15 = new System.Windows.Forms.Label();
            this.lbl20 = new System.Windows.Forms.Label();
            this.lbl25 = new System.Windows.Forms.Label();
            this.lbl30 = new System.Windows.Forms.Label();

            ((System.ComponentModel.ISupportInitialize)(this.tbSeconds)).BeginInit();
            this.SuspendLayout();
            // 
            // chkContinuous
            // 
            this.chkContinuous.AutoSize = true;
            this.chkContinuous.ForeColor = System.Drawing.Color.Gainsboro;
            this.chkContinuous.Location = new System.Drawing.Point(16, 16);
            this.chkContinuous.Name = "chkContinuous";
            this.chkContinuous.Size = new System.Drawing.Size(162, 20);
            this.chkContinuous.TabIndex = 0;
            this.chkContinuous.Text = "Continuous (auto-stop)";
            // 
            // tbSeconds
            // 
            this.tbSeconds.Location = new System.Drawing.Point(16, 44);
            this.tbSeconds.Maximum = 30;
            this.tbSeconds.Minimum = 1;
            this.tbSeconds.Name = "tbSeconds";
            this.tbSeconds.Size = new System.Drawing.Size(280, 56);
            this.tbSeconds.TabIndex = 1;
            this.tbSeconds.Value = 15;
            // 
            // lblSeconds
            // 
            this.lblSeconds.AutoSize = true;
            this.lblSeconds.ForeColor = System.Drawing.Color.Gainsboro;
            this.lblSeconds.Location = new System.Drawing.Point(32, 84);
            this.lblSeconds.Name = "lblSeconds";
            this.lblSeconds.Size = new System.Drawing.Size(90, 16);
            this.lblSeconds.TabIndex = 2;
            this.lblSeconds.Text = "Listen for: 15 s";
            // 
            // chkSingle
            // 
            this.chkSingle.AutoSize = true;
            this.chkSingle.ForeColor = System.Drawing.Color.Gainsboro;
            this.chkSingle.Location = new System.Drawing.Point(16, 120);
            this.chkSingle.Name = "chkSingle";
            this.chkSingle.Size = new System.Drawing.Size(258, 20);
            this.chkSingle.TabIndex = 3;
            this.chkSingle.Text = "Single utterance (≈3s silence, 30s max)";

            // 
            // tlpTicks  (numeric guides under the slider)
            // 
            this.tlpTicks.ColumnCount = 7;
            this.tlpTicks.RowCount = 1;
            this.tlpTicks.Location = new System.Drawing.Point(16, 104);  // under the slider
            this.tlpTicks.Name = "tlpTicks";
            this.tlpTicks.Size = new System.Drawing.Size(280, 18);       // match slider width
            this.tlpTicks.TabIndex = 4;
            this.tlpTicks.BackColor = System.Drawing.Color.Transparent;

            // evenly-spaced columns
            this.tlpTicks.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / 6F));
            this.tlpTicks.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / 6F));
            this.tlpTicks.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / 6F));
            this.tlpTicks.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / 6F));
            this.tlpTicks.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / 6F));
            this.tlpTicks.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / 6F));
            this.tlpTicks.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / 6F));
            this.tlpTicks.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 
            // labels (1, 5, 10, 15, 20, 25, 30)
            // 
            var labelFont = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            this.lbl1.Text = "1"; this.lbl1.ForeColor = Color.Gainsboro; this.lbl1.Dock = DockStyle.Fill; this.lbl1.TextAlign = ContentAlignment.TopCenter; this.lbl1.Font = labelFont;
            this.lbl5.Text = "5"; this.lbl5.ForeColor = Color.Gainsboro; this.lbl5.Dock = DockStyle.Fill; this.lbl5.TextAlign = ContentAlignment.TopCenter; this.lbl5.Font = labelFont;
            this.lbl10.Text = "10"; this.lbl10.ForeColor = Color.Gainsboro; this.lbl10.Dock = DockStyle.Fill; this.lbl10.TextAlign = ContentAlignment.TopCenter; this.lbl10.Font = labelFont;
            this.lbl15.Text = "15"; this.lbl15.ForeColor = Color.Gainsboro; this.lbl15.Dock = DockStyle.Fill; this.lbl15.TextAlign = ContentAlignment.TopCenter; this.lbl15.Font = labelFont;
            this.lbl20.Text = "20"; this.lbl20.ForeColor = Color.Gainsboro; this.lbl20.Dock = DockStyle.Fill; this.lbl20.TextAlign = ContentAlignment.TopCenter; this.lbl20.Font = labelFont;
            this.lbl25.Text = "25"; this.lbl25.ForeColor = Color.Gainsboro; this.lbl25.Dock = DockStyle.Fill; this.lbl25.TextAlign = ContentAlignment.TopCenter; this.lbl25.Font = labelFont;
            this.lbl30.Text = "30"; this.lbl30.ForeColor = Color.Gainsboro; this.lbl30.Dock = DockStyle.Fill; this.lbl30.TextAlign = ContentAlignment.TopCenter; this.lbl30.Font = labelFont;

            this.tlpTicks.Controls.Add(this.lbl1, 0, 0);
            this.tlpTicks.Controls.Add(this.lbl5, 1, 0);
            this.tlpTicks.Controls.Add(this.lbl10, 2, 0);
            this.tlpTicks.Controls.Add(this.lbl15, 3, 0);
            this.tlpTicks.Controls.Add(this.lbl20, 4, 0);
            this.tlpTicks.Controls.Add(this.lbl25, 5, 0);
            this.tlpTicks.Controls.Add(this.lbl30, 6, 0);

            // 
            // SettingsPane
            // 
            this.BackColor = System.Drawing.Color.DimGray;
            this.Controls.Add(this.chkContinuous);
            this.Controls.Add(this.tbSeconds);
            this.Controls.Add(this.lblSeconds);
            this.Controls.Add(this.tlpTicks);
            this.Controls.Add(this.chkSingle);
            this.Name = "SettingsPane";
            this.Size = new System.Drawing.Size(340, 220);
            ((System.ComponentModel.ISupportInitialize)(this.tbSeconds)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
