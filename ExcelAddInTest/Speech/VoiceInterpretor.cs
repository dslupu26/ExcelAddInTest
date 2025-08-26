using Microsoft.CognitiveServices.Speech;
using System;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest
{
    public class VoiceInterpretor
    {
        private SpeechRecognizer recognizer;
        string speechKey = Config.SpeechKey;
        string speechRegion = Config.SpeechRegion;

        private CluService _clu; //variabila privata pentru serviciul CLU
        public VoiceInterpretor(CluService clu)
        {
            _clu = clu; //injectam clu prin constructor
        }

        public async Task VoiceToExcelAsync()
        {
            try
            {
                Globals.ThisAddIn.AppendToPane("[Speech] starting…");

                var config = SpeechConfig.FromSubscription(Config.SpeechKey, Config.SpeechRegion);
                config.SpeechRecognitionLanguage = Config.SpeechLanguage;   // "en-US"

                recognizer = new SpeechRecognizer(config /*, audio*/);

                // ——— LOG ALL EVENTS ———
                recognizer.SessionStarted += (s, e) =>
                    Globals.ThisAddIn.AppendToPane("[Speech] SessionStarted");
                recognizer.SessionStopped += (s, e) =>
                    Globals.ThisAddIn.AppendToPane("[Speech] SessionStopped");

                recognizer.SpeechStartDetected += (s, e) =>
                    Globals.ThisAddIn.AppendToPane("[Speech] SpeechStartDetected");
                recognizer.SpeechEndDetected += (s, e) =>
                    Globals.ThisAddIn.AppendToPane("[Speech] SpeechEndDetected");

                recognizer.Recognizing += (s, e) =>
                {
                    if (e.Result.Reason == ResultReason.RecognizingSpeech)
                        Globals.ThisAddIn.AppendToPane("[Speech] Recognizing: " + e.Result.Text);
                };

                recognizer.Recognized += async (s, e) =>
                {
                    Globals.ThisAddIn.AppendToPane("[Speech] Recognized reason: " + e.Result.Reason);
                    if (e.Result.Reason == ResultReason.RecognizedSpeech)
                    {
                        string text = e.Result.Text;
                        Globals.ThisAddIn.AppendToPane("[Speech] Final: " + text);

                        // (opțional) trimit la CLU
                        try
                        {
                            var nlu = await _clu.AnalyzeAsync(text);
                            Globals.ThisAddIn.AppendToPane("[CLU RAW]\r\n" + nlu.RawJson);
                            Globals.ThisAddIn.AppendToPane("[CLU] TopIntent: " + nlu.TopIntent);
                            foreach (var ent in nlu.Entities)
                                Globals.ThisAddIn.AppendToPane(" - " + ent.Category + ": \"" + ent.Text + "\"");
                        }
                        catch (Exception exClu)
                        {
                            Globals.ThisAddIn.AppendToPane("[CLU] ERROR: " + exClu.Message);
                        }
                    }
                    else if (e.Result.Reason == ResultReason.NoMatch)
                    {
                        Globals.ThisAddIn.AppendToPane("[Speech] NoMatch");
                    }
                };

                recognizer.Canceled += (s, e) =>
                {
                    Globals.ThisAddIn.AppendToPane("[Speech] Canceled: " + e.Reason +
                        (e.Reason == CancellationReason.Error
                            ? " | " + e.ErrorCode + " | " + e.ErrorDetails
                            : ""));
                };

                await recognizer.StartContinuousRecognitionAsync();
                Globals.ThisAddIn.AppendToPane("[Speech] Started (speak now)");
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.AppendToPane("[Speech] START ERROR: " + ex.Message);
            }
        }

        public async Task VoiceToExcelStopAync()
        {
            try
            {
                if (recognizer != null)
                {
                    Globals.ThisAddIn.AppendToPane("[Speech] stopping…");
                    await recognizer.StopContinuousRecognitionAsync();
                    recognizer.Dispose();
                    recognizer = null;
                    Globals.ThisAddIn.AppendToPane("[Speech] stopped.");
                }
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.AppendToPane("[Speech] STOP ERROR: " + ex.Message);
            }
        }
    }
}
