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
        private CustomTaskPane _settingsPane;
        private VoiceListenOptions vlo = new VoiceListenOptions() // Default setting
        {
            Mode = ListenMode.Continuous,
            AutoStopAfter = null,
            MaxDuration = TimeSpan.FromSeconds(30) // nu asculta mai mult de 30 de secunde decat daca il schimbam noi

            // Set By Default
            // Language = Config.SpeechLanguage,
            // InitialSilenceTimeoutMs = 5000,
            // EndSilenceTimeoutMs = 3000
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

        public void SetSettingsPane(CustomTaskPane settingsPane)
        {
            _settingsPane = settingsPane;
        }

        private async void StartRecording(object sender, EventArgs e)
        {
            await _voice.StartAsync(vlo);
        }

        public void SetOptionsRef(VoiceListenOptions opts) => this.vlo = opts;

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

        

        private void UserControlPane_Load(object sender, EventArgs e)
        {

        }

        private void settingsBtn_Click(object sender, EventArgs e)
        {
            if (_settingsPane != null)
            {
                _settingsPane.Visible = !_settingsPane.Visible;
                return;
            };
        }
    }
}
