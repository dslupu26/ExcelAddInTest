using Microsoft.CognitiveServices.Speech;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest
{
    public partial class Ribbon1
    {
        private VoiceInterpreter _voice;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        { 

        }

/*        private void btnSum_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            double suma = (double)Globals.ThisAddIn.Application.WorksheetFunction.Sum(ws.Range["A1", "A4"]);
            ws.Range["B1"].Value = suma;
        }*/

        private async void startRecord_Click(object sender, RibbonControlEventArgs e)
        {
            await _voice.VoiceToExcelAsync();

        }

        private void SpeechBox_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }
        private void btnShowPane_Click(object sender, RibbonControlEventArgs e)
        {

            Globals.ThisAddIn.AppendToPane("Pane test: hello from Ribbon.");
        }

        public async void stopRecord_Click(object sender, RibbonControlEventArgs e)
        {

            _voice = Globals.ThisAddIn.GetOrCreateVoice();
            await _voice.VoiceToExcelStopAync();
        }

    }
}
