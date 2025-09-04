using ExcelAddInTest;
using ExcelAddInTest.ExcelApi;
using ExcelAddInTest.Logging;
using ExcelAddInTest.Nlu;
using ExcelAddInTest.Utils;
using Microsoft.CognitiveServices.Speech;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

public class VoiceInterpreter
{
    private SpeechRecognizer recognizer;
    private readonly INlu _clu;
    private readonly ICommandExecutor _exec;
    private readonly ILogger _log;
    private readonly IIntentRouter intentRouter;

    private CancellationTokenSource _cts;
    private bool _isListening;
    private VoiceListenOptions _opts;

    private EntityDistributor _ent;


    public VoiceInterpreter(CluService clu, ICommandExecutor excel, ExcelAddInTest.Logging.ILogger log, EntityDistributor ent, IIntentRouter intentRouter)
    {
        _clu = clu;
        _exec = excel;
        _log = log ?? throw new ArgumentNullException(nameof(log));
        _ent = ent;
        this.intentRouter = intentRouter;
    }

    public async Task StartAsync(VoiceListenOptions opts)
    {
        if (_isListening) return;
        _isListening = true;
        _opts = opts ?? new VoiceListenOptions();
        _cts = new CancellationTokenSource();

        _log.Info("starting…");

        var config = SpeechConfig.FromSubscription(Config.SpeechKey, Config.SpeechRegion);
        _ = config ?? throw new InvalidOperationException("Speech configuration failed.");
        config.SpeechRecognitionLanguage = _opts.Language;

        // Silence timeouts – control how fast an utterance is considered finished
        config.SetProperty(PropertyId.SpeechServiceConnection_InitialSilenceTimeoutMs,
                           _opts.InitialSilenceTimeoutMs.ToString());
        config.SetProperty(PropertyId.SpeechServiceConnection_EndSilenceTimeoutMs,
                           _opts.EndSilenceTimeoutMs.ToString());

        recognizer = new SpeechRecognizer(config); // (optionally) pass AudioConfig for specific mic
        WireEvents();

        if (_opts.Mode == ListenMode.SingleUtterance)
        {
            try
            {
                var result = await recognizer.RecognizeOnceAsync().WithCancellation(_cts.Token);
                await HandleResultAsync(result);
            }
            catch (OperationCanceledException)
            {
                _log.Warn("Single-utterance canceled.");
            }
            finally
            {
                await StopAsync();
            }
            return;
        }

        // Continuous
        await recognizer.StartContinuousRecognitionAsync();
        _log.Info("Started (speak now)");

        // Auto-stop after a duration (e.g., 15/30s)
        if (_opts.AutoStopAfter.HasValue)
            _ = GuardStopAfter(_opts.AutoStopAfter.Value, "[AutoStopAfter]");

        // Hard cap (until you press Stop, but max 30s)
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
                _log.Info("stopping…");
                try { await recognizer.StopContinuousRecognitionAsync(); } catch { /* already stopped */ }
                recognizer.Dispose();
                recognizer = null;
            }
            _log.Info("stopped.");
        }
        catch (Exception ex)
        {
            _log.Error("STOP ERROR", ex);
        }
    }

    // —— Helpers ——
    private async Task GuardStopAfter(TimeSpan delay, string tag)
    {
        try
        {
            await Task.Delay(delay, _cts.Token);
            _log.Info($"{tag} elapsed → stopping.");
            await StopAsync();
        }
        catch (TaskCanceledException) { /* ignore */ }
    }

    private void WireEvents()
    {
        recognizer.SessionStarted += (s, e) => _log.Info("SessionStarted");
        recognizer.SessionStopped += (s, e) => { _log.Info("SessionStopped"); };
        recognizer.SpeechStartDetected += (s, e) => _log.Info("SpeechStartDetected");
        recognizer.SpeechEndDetected += (s, e) => _log.Info("SpeechEndDetected");

        recognizer.Recognizing += (s, e) =>
        {
            if (e.Result.Reason == ResultReason.RecognizingSpeech)
                _log.Info("Recognizing: " + e.Result.Text);
        };

        recognizer.Recognized += async (s, e) =>
        {
            _log.Info("Recognized reason: " + e.Result.Reason);
            if (e.Result.Reason == ResultReason.RecognizedSpeech)
            {
                await HandleResultAsync(e.Result);
            }
        };

        recognizer.Canceled += (s, e) =>
        {
            if (e.Reason == CancellationReason.Error)
                _log.Error($"Canceled: {e.Reason} | {e.ErrorCode} | {e.ErrorDetails}");
            else
                _log.Warn($"Canceled: {e.Reason}");
        };
    }


    // REMOVE THIS
    // THIS IS HERE FOR ""BOOKMARKING"" PURPOSES
    // SO YOU DONT HAVE TO SEARCH FOR THE RAW CLU INPUT ANYMORE

    /// <summary>
    /// Process the final recognized text, calls CLU for the intent and routes the command, then executes it
    /// </summary>
    /// <param name="result"></param>
    /// <returns></returns>
    private async Task HandleResultAsync(SpeechRecognitionResult result)
    {
        var text = result.Text?.Trim();


        // this line here has the FINAL result
        _log.Info("Final: " + text);


        //If the speech service is still listening and identifies no text,
        //(e.g. the person does not speak or the speech is not recognized),
        //we return without doing anything, so that we do not call CLU with empty text.
        if (string.IsNullOrWhiteSpace(text))
        {
            _log.Warn("No speech recognized.");
            return;
        }
        try
        {
            var nlu = await _clu.AnalyzeAsync(text);
            _log.Raw("[CLU RAW]\r\n" + nlu.RawJson);
            _log.Info("[CLU] TopIntent: " + nlu.TopIntent);


            // nlu.Entities has the cells; lowkey no need to parse them. again.

            foreach (var ent in nlu.Entities)
                _log.Info($" - {ent.Category}: \"{ent.Text}\"");


            // not done yet
            _ent.ListMaker(result);
        }
        catch (Exception exClu)
        {
            _log.Error("[CLU] ERROR", exClu);
        }
    }
}
