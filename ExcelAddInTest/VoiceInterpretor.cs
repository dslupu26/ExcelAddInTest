using Microsoft.CognitiveServices.Speech;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest
{
    public class VoiceInterpretor
    {
        private SpeechRecognizer recognizer;
        string speechKey = "8yZ5oIsWR4zIAurWr2vz5ORVw6gqNLD8PlgnaqDJpzNyJwUhR5XAJQQJ99BHAC5RqLJXJ3w3AAAYACOGHeQu";
        string speechRegion = "westeurope";


        public async Task VoiceToExcelAsync()
        {
            var config = SpeechConfig.FromSubscription(speechKey, speechRegion);
            config.SpeechRecognitionLanguage = "en-US";

            recognizer = new SpeechRecognizer(config);

            recognizer.Recognizing += (s, e) =>
            {
                Globals.Ribbons.Ribbon1.speechBox.Text = e.Result.Text;
            };

            recognizer.Recognized += (s, e) =>
            {
                Globals.Ribbons.Ribbon1.speechBox.Text = e.Result.Text;
            };

            recognizer.Canceled += (s, e) =>
            {
                System.Windows.Forms.MessageBox.Show($"Canceled: {e.Reason}");
            };

            recognizer.SessionStopped += (s, e) =>
            {
                System.Windows.Forms.MessageBox.Show("Speech not recognized. Please try again.");
            };

            await recognizer.StartContinuousRecognitionAsync();


            // System.Windows.Forms.MessageBox.Show("Please speak your command for Excel.");
            /*var result = await recognizer.RecognizeOnceAsync();

            if (result.Reason == ResultReason.RecognizedSpeech)
            {
                string text = result.Text;
                // System.Windows.Forms.MessageBox.Show($"You said: {text}");

                Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                // ws.Range["A1"].Value = text;
                Globals.Ribbons.Ribbon1.speechBox.Text = text;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Speech not recognized. Please try again.");
            }*/
        }

        public async Task VoiceToExcelStopAync()
        {
            if (recognizer != null)
            {
                await recognizer.StopContinuousRecognitionAsync();
                recognizer.Dispose();
                recognizer = null;
            }
        }
    }
}
