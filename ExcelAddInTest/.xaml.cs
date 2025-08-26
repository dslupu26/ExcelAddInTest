using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelAddInTest
{
    public partial class SpeechPanel : UserControl
    {
        public SpeechPanel()
        {
            InitializeComponent();
        }

        public void AppendText(string text)
        {
            speechBox.Dispatcher.Invoke(() =>
            {
                speechBox.AppendText(text + " ");
                speechBox.ScrollToEnd();
            });
        }
    }
}
