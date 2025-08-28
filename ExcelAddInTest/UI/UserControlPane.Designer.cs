namespace ExcelAddInTest.UserInterface
{
    partial class UserControlPane
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
            this.InputBox = new System.Windows.Forms.TextBox();
            this.Input = new System.Windows.Forms.Label();
            this.OutputBox = new System.Windows.Forms.TextBox();
            this.Output = new System.Windows.Forms.Label();
            this.nameLabel = new System.Windows.Forms.Label();
            this.startBtn = new System.Windows.Forms.Button();
            this.stopBtn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // InputBox
            // 
            this.InputBox.Location = new System.Drawing.Point(39, 90);
            this.InputBox.Multiline = true;
            this.InputBox.Name = "InputBox";
            this.InputBox.ReadOnly = true;
            this.InputBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.InputBox.Size = new System.Drawing.Size(500, 106);
            this.InputBox.TabIndex = 0;
            // 
            // Input
            // 
            this.Input.AutoSize = true;
            this.Input.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Input.ForeColor = System.Drawing.Color.Gainsboro;
            this.Input.Location = new System.Drawing.Point(35, 65);
            this.Input.Name = "Input";
            this.Input.Size = new System.Drawing.Size(54, 22);
            this.Input.TabIndex = 1;
            this.Input.Text = "Input";
            // 
            // OutputBox
            // 
            this.OutputBox.Location = new System.Drawing.Point(39, 242);
            this.OutputBox.Multiline = true;
            this.OutputBox.Name = "OutputBox";
            this.OutputBox.ReadOnly = true;
            this.OutputBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.OutputBox.Size = new System.Drawing.Size(500, 106);
            this.OutputBox.TabIndex = 2;
            // 
            // Output
            // 
            this.Output.AutoSize = true;
            this.Output.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Output.ForeColor = System.Drawing.Color.Gainsboro;
            this.Output.Location = new System.Drawing.Point(35, 217);
            this.Output.Name = "Output";
            this.Output.Size = new System.Drawing.Size(70, 22);
            this.Output.TabIndex = 3;
            this.Output.Text = "Output";
            // 
            // nameLabel
            // 
            this.nameLabel.AutoSize = true;
            this.nameLabel.BackColor = System.Drawing.Color.DimGray;
            this.nameLabel.Font = new System.Drawing.Font("Times New Roman", 16.2F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nameLabel.ForeColor = System.Drawing.Color.Gainsboro;
            this.nameLabel.Location = new System.Drawing.Point(135, 17);
            this.nameLabel.Name = "nameLabel";
            this.nameLabel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.nameLabel.Size = new System.Drawing.Size(309, 32);
            this.nameLabel.TabIndex = 4;
            this.nameLabel.Text = "EXCEL VOICE ADD-IN";
            this.nameLabel.Click += new System.EventHandler(this.nameLabel_Click);
            // 
            // startBtn
            // 
            this.startBtn.AutoSize = true;
            this.startBtn.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.startBtn.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.startBtn.Location = new System.Drawing.Point(39, 377);
            this.startBtn.Name = "startBtn";
            this.startBtn.Size = new System.Drawing.Size(146, 54);
            this.startBtn.TabIndex = 5;
            this.startBtn.Text = "Record";
            this.startBtn.UseVisualStyleBackColor = false;
            this.startBtn.Click += new System.EventHandler(this.StartRecording);
            // 
            // stopBtn
            // 
            this.stopBtn.AutoSize = true;
            this.stopBtn.BackColor = System.Drawing.Color.Salmon;
            this.stopBtn.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.stopBtn.Location = new System.Drawing.Point(191, 377);
            this.stopBtn.Name = "stopBtn";
            this.stopBtn.Size = new System.Drawing.Size(146, 54);
            this.stopBtn.TabIndex = 6;
            this.stopBtn.Text = "Stop Recording";
            this.stopBtn.UseVisualStyleBackColor = false;
            this.stopBtn.Click += new System.EventHandler(this.StopRecording);
            // 
            // button1
            // 
            this.button1.AutoSize = true;
            this.button1.BackColor = System.Drawing.Color.Olive;
            this.button1.Font = new System.Drawing.Font("French Script MT", 19.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(385, 377);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(154, 54);
            this.button1.TabIndex = 7;
            this.button1.Text = "Debug";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.DebugButton_Click);
            // 
            // UserControlPane
            // 
            this.AccessibleName = "UserC";
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DimGray;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.stopBtn);
            this.Controls.Add(this.startBtn);
            this.Controls.Add(this.nameLabel);
            this.Controls.Add(this.Output);
            this.Controls.Add(this.OutputBox);
            this.Controls.Add(this.Input);
            this.Controls.Add(this.InputBox);
            this.Name = "UserControlPane";
            this.Size = new System.Drawing.Size(580, 560);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox InputBox;
        private System.Windows.Forms.Label Input;
        private System.Windows.Forms.TextBox OutputBox;
        private System.Windows.Forms.Label Output;
        private System.Windows.Forms.Label nameLabel;
        private System.Windows.Forms.Button startBtn;
        private System.Windows.Forms.Button stopBtn;
        private System.Windows.Forms.Button button1;
    }
}
