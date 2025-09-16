using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddInTest.UserInterface
{
    public partial class UserControlPane : System.Windows.Forms.UserControl
    {
        private VoiceInterpreter _voice;
        private CustomTaskPane _debugPane;
        private VoiceListenOptions vlo = new VoiceListenOptions()
        {
            Mode = ListenMode.Continuous,
            AutoStopAfter = null,
            MaxDuration = TimeSpan.FromSeconds(30),
            Language = Config.SpeechLanguage,
            InitialSilenceTimeoutMs = 5000,
            EndSilenceTimeoutMs = 2000
        };
        public UserControlPane()
        {
            InitializeComponent();
        }

        private void nameLabel_Click(object sender, EventArgs e)
        {

        }

        public void SetVoiceInterpreter(VoiceInterpreter voice)
        {
            _voice = voice;
        }
        public void SetDebugPane(CustomTaskPane debugPane)
        {
            _debugPane = debugPane;
        }

        private async void StartRecording(object sender, EventArgs e)
        {
            await _voice.StartAsync(vlo);
        }

        private async void StopRecording(object sender, EventArgs e)
        {
            await _voice.StopAsync();
        }

        private void DebugButton_Click(object sender, EventArgs e)
        {
            if (_debugPane != null)
            {
                _debugPane.Visible = !_debugPane.Visible;
                return;
            };
        }

        internal void AppendInput(string v)
        {   if(InvokeRequired)
            {
                this.Invoke(new Action<string>(AppendInput), v);
                return;
            }
            this.InputBox.AppendText(v + Environment.NewLine);
        }

        internal void AppendOutput(string v)
        {
            if(InvokeRequired)
            {
                this.Invoke(new Action<string>(AppendOutput), v);
                return;
            }
            this.OutputBox.AppendText(v + Environment.NewLine);
        }

        private void UserControlPane_Load(object sender, EventArgs e)
        {

        }
    }
}
