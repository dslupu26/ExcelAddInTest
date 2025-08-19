namespace ExcelAddInTest
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnSum = this.Factory.CreateRibbonButton();
            this.startRecord = this.Factory.CreateRibbonButton();
            this.stopRecord = this.Factory.CreateRibbonButton();
            this.speechBox = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSum);
            this.group1.Items.Add(this.startRecord);
            this.group1.Items.Add(this.stopRecord);
            this.group1.Items.Add(this.speechBox);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnSum
            // 
            this.btnSum.Label = "btnSum";
            this.btnSum.Name = "btnSum";
            this.btnSum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSum_Click);
            // 
            // startRecord
            // 
            this.startRecord.Label = "startRecord";
            this.startRecord.Name = "startRecord";
            this.startRecord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.startRecord_Click);
            // 
            // stopRecord
            // 
            this.stopRecord.Label = "stopRecord";
            this.stopRecord.Name = "stopRecord";
            this.stopRecord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.stopRecord_Click);
            // 
            // speechBox
            // 
            this.speechBox.Label = "speechBox";
            this.speechBox.Name = "speechBox";
            this.speechBox.Text = null;
            this.speechBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.speechBox_TextChanged);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSum;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton startRecord;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox speechBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton stopRecord;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
