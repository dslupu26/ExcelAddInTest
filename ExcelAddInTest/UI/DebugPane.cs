using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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
