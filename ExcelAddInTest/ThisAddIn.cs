using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Net;

namespace ExcelAddInTest
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _cluPane;
        private CluOutputPane _cluControl;

        public Microsoft.Office.Tools.CustomTaskPane CluPane => _cluPane;
        public CluOutputPane CluControl => _cluControl;

        public VoiceInterpretor Voice { get; private set; }

        public static SynchronizationContext UiContext { get; private set; }       

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            MessageBox.Show("ThisAddIn_Startup called");
            System.Net.ServicePointManager.SecurityProtocol =System.Net.SecurityProtocolType.Tls12;
            try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)12288; } catch { }
            EnsureCluPane();           
            _cluPane.Visible = true;
            GetOrCreateVoice();
        }
        public void EnsureCluPane()
        {
            if (_cluPane != null) return;

            _cluControl = new CluOutputPane(); // UserControl cu TextBox multiline, Dock=Fill
            _cluPane = this.CustomTaskPanes.Add(_cluControl, "CLU Output");

            // poziție & dimensiuni ca să fie sigur vizibil
            _cluPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _cluPane.Visible = true;
        }



        public VoiceInterpretor GetOrCreateVoice()
        {
            if (Voice == null)
            { 
                var clu = new CluService( Config.CluEndpoint, Config.CluKey, Config.CluProjectName,
                    Config.CluDeployment);
                Voice = new VoiceInterpretor(clu);
            }
            return Voice;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        public void AppendToPane(string text)
        {
            EnsureCluPane();
            if (_cluControl == null) return;
            _cluControl.AppendText(text);
        }

        public void AppendCluLog(ExcelAddInTest.NluResult nlu)
        {
            if (nlu == null) return;

            var sb = new StringBuilder();
            sb.AppendLine("TopIntent: " + (nlu.TopIntent ?? "None"));

            if (nlu.Entities != null && nlu.Entities.Count > 0)
            {
                sb.AppendLine("Entities:");
                foreach (var ent in nlu.Entities)
                    sb.AppendLine(" - " + ent.Category + ": \"" + ent.Text + "\"");
            }
            else
            {
                sb.AppendLine("Entities: (none)");
            }

            // opțional: vezi și JSON-ul brut
            // sb.AppendLine();
            // sb.AppendLine(nlu.RawJson);

            // trimitem textul în panoul tău CluOutputPane
            _cluControl.AppendText(sb.ToString());
        }

    }
}
