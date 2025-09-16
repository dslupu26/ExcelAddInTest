using System;
using System.Windows.Forms;

namespace ExcelAddInTest
{
    public partial class DebugPane : UserControl
    {
        public DebugPane()
        {
            InitializeComponent();
        }

        public void AppendText(string text)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(AppendText), text);
                return;
            }
            CluOutputBox.AppendText(text + Environment.NewLine);
        }
    }
}
