using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.CognitiveServices.Speech;

namespace ExcelAddInTest
{
    public partial class Ribbon1
    {
        private VoiceInterpretor voiceInterpretor = new VoiceInterpretor();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSum_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            double suma = (double)Globals.ThisAddIn.Application.WorksheetFunction.Sum(ws.Range["A1", "A4"]);
            ws.Range["B1"].Value = suma;
        }

        private async void startRecord_Click(object sender, RibbonControlEventArgs e)
        {
            await voiceInterpretor.VoiceToExcelAsync();

        }

        private void speechBox_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        public async void stopRecord_Click(object sender, RibbonControlEventArgs e)
        {
            await voiceInterpretor.VoiceToExcelStopAync();
        }
    }
}
