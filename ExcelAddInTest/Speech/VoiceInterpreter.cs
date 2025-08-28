using Microsoft.CognitiveServices.Speech;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelAddInTest.Utils;

public class VoiceInterpretor
{
    private SpeechRecognizer recognizer;
    private readonly CluService _clu;

    private CancellationTokenSource _cts;
    private bool _isListening;
    private VoiceListenOptions _opts;

    public VoiceInterpretor(CluService clu) => _clu = clu;

    public async Task StartAsync(VoiceListenOptions opts)
    {
        if (_isListening) return;
        _isListening = true;
        _opts = opts ?? new VoiceListenOptions();
        _cts = new CancellationTokenSource();

        Globals.ThisAddIn.AppendToPane("[Speech] starting…");

        var config = SpeechConfig.FromSubscription(Config.SpeechKey, Config.SpeechRegion);
        config.SpeechRecognitionLanguage = _opts.Language;

        // Timpii de tăcere (silence) – controlează cât de repede „taie” fraza
        config.SetProperty(PropertyId.SpeechServiceConnection_InitialSilenceTimeoutMs,
                           _opts.InitialSilenceTimeoutMs.ToString());
        config.SetProperty(PropertyId.SpeechServiceConnection_EndSilenceTimeoutMs,
                           _opts.EndSilenceTimeoutMs.ToString());

        recognizer = new SpeechRecognizer(config); // poți trece și AudioConfig pt. device selectat

        WireEvents(); // atașează evenimentele tale de log + CLU

        if (_opts.Mode == ListenMode.SingleUtterance)
        {
            try
            {
                var result = await recognizer.RecognizeOnceAsync().WithCancellation(_cts.Token);
                await HandleResultAsync(result);
            }
            catch (OperationCanceledException)
            {
                Globals.ThisAddIn.AppendToPane("[Speech] Single-utterance canceled.");
            }
            finally
            {
                await StopAsync();
            }
            return;
        }

        // Continuous
        await recognizer.StartContinuousRecognitionAsync();
        Globals.ThisAddIn.AppendToPane("[Speech] Started (speak now)");

        // Auto-stop după o durată (ex. 15/30s)
        if (_opts.AutoStopAfter.HasValue)
            _ = GuardStopAfter(_opts.AutoStopAfter.Value, "[AutoStopAfter]");

        // Hard cap (ex. până apeși Stop, dar max 30s)
        if (_opts.MaxDuration.HasValue)
            _ = GuardStopAfter(_opts.MaxDuration.Value, "[MaxDuration]");
    }

    public async Task StopAsync()
    {
        if (!_isListening) return;
        _isListening = false;

        try
        {
            _cts?.Cancel();
            if (recognizer != null)
            {
                Globals.ThisAddIn.AppendToPane("[Speech] stopping…");
                try { await recognizer.StopContinuousRecognitionAsync(); } catch { /* poate fi deja oprit */ }
                recognizer.Dispose();
                recognizer = null;
            }
            Globals.ThisAddIn.AppendToPane("[Speech] stopped.");
        }
        catch (Exception ex)
        {
            Globals.ThisAddIn.AppendToPane("[Speech] STOP ERROR: " + ex.Message);
        }
    }

    // —— Helpers ——
    private async Task GuardStopAfter(TimeSpan delay, string tag)
    {
        try
        {
            await Task.Delay(delay, _cts.Token);
            Globals.ThisAddIn.AppendToPane($"{tag} elapsed → stopping.");
            await StopAsync();
        }
        catch (TaskCanceledException) { /* ignoră */ }
    }

    private void WireEvents()
    {
        recognizer.SessionStarted += (s, e) => Globals.ThisAddIn.AppendToPane("[Speech] SessionStarted");
        recognizer.SessionStopped += (s, e) => Globals.ThisAddIn.AppendToPane("[Speech] SessionStopped");
        recognizer.SpeechStartDetected += (s, e) => Globals.ThisAddIn.AppendToPane("[Speech] SpeechStartDetected");
        recognizer.SpeechEndDetected += (s, e) => Globals.ThisAddIn.AppendToPane("[Speech] SpeechEndDetected");

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
                await HandleResultAsync(e.Result);

                // dacă vrei „oprește-te după prima frază” DAR rămâi pe continuous:
                // if (_opts.AutoStopAfter == null && _opts.MaxDuration == null && someFlag)
                //     await StopAsync();
            }
        };

        recognizer.Canceled += (s, e) =>
        {
            Globals.ThisAddIn.AppendToPane("[Speech] Canceled: " + e.Reason +
                (e.Reason == CancellationReason.Error ? " | " + e.ErrorCode + " | " + e.ErrorDetails : ""));
        };
    }

    private async Task HandleResultAsync(SpeechRecognitionResult result)
    {
        var text = result.Text?.Trim();
        Globals.ThisAddIn.AppendToPane("[Speech] Final: " + text);

        try
        {
            var nlu = await _clu.AnalyzeAsync(text);
            Globals.ThisAddIn.AppendToPane("[CLU RAW]\r\n" + nlu.RawJson);
            Globals.ThisAddIn.AppendToPane("[CLU] TopIntent: " + nlu.TopIntent);
            foreach (var ent in nlu.Entities)
                Globals.ThisAddIn.AppendToPane($" - {ent.Category}: \"{ent.Text}\"");
        }
        catch (Exception exClu)
        {
            Globals.ThisAddIn.AppendToPane("[CLU] ERROR: " + exClu.Message);
        }
    }
}
