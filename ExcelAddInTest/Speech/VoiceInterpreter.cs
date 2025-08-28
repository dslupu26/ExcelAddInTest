using ExcelAddInTest;
using ExcelAddInTest.ExcelApi;
using ExcelAddInTest.Utils;
using Microsoft.CognitiveServices.Speech;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

public class VoiceInterpreter
{
    private SpeechRecognizer recognizer;
    private readonly CluService _clu;
    private readonly IExcelActions _excel;
    private readonly ExcelAddInTest.Logging.ILogger _log;

    private CancellationTokenSource _cts;
    private bool _isListening;
    private VoiceListenOptions _opts;

    public VoiceInterpreter(CluService clu, IExcelActions excel, ExcelAddInTest.Logging.ILogger log)
    {
        _clu = clu;
        _excel = excel;
        _log = log ?? throw new ArgumentNullException(nameof(log));
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
        recognizer.SessionStopped += (s, e) => _log.Info("SessionStopped");
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

    private async Task HandleResultAsync(SpeechRecognitionResult result)
    {
        var text = result.Text?.Trim();
        _log.Info("Final: " + text);

        try
        {
            var nlu = await _clu.AnalyzeAsync(text);
            _log.Raw("[CLU RAW]\r\n" + nlu.RawJson);
            _log.Info("[CLU] TopIntent: " + nlu.TopIntent);
            foreach (var ent in nlu.Entities)
                _log.Info($" - {ent.Category}: \"{ent.Text}\"");
        }
        catch (Exception exClu)
        {
            _log.Error("[CLU] ERROR", exClu);
        }
    }
}
