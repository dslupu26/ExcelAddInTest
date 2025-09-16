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

public class VoiceInterpreter
{
    private SpeechRecognizer recognizer;
    private readonly INlu _clu;
    private readonly ICommandExecutor _exec;
    private readonly ILogger _log;
    private readonly IIntentRouter _intentRouter;

    private CancellationTokenSource _cts;          // sesiunea curentă
    private CancellationTokenSource _deadlineCts;  // timerul de auto-stop (rearmabil)
    private bool _isListening;
    private VoiceListenOptions _opts = new VoiceListenOptions();

    private readonly EntityDistributor _ent;

    public VoiceInterpreter(
        INlu clu,                           // <- folosește interfața aici
        ICommandExecutor exec,
        ILogger log,
        EntityDistributor ent,
        IIntentRouter intentRouter)

    public VoiceInterpreter(CluService clu, ICommandExecutor excel, ILogger log, EntityDistributor ent, IIntentRouter intentRouter)
    {
        _clu = clu ?? throw new ArgumentNullException(nameof(clu));
        _exec = exec;
        _log = log ?? throw new ArgumentNullException(nameof(log));
        _ent = ent;
        _intentRouter = intentRouter;
    }

    public async Task StartAsync(VoiceListenOptions opts)
    {
        if (_isListening) return;
        _isListening = true;
        _opts = opts ?? new VoiceListenOptions();
        _cts = new CancellationTokenSource();

        _log.Info("starting…");

        var config = SpeechConfig.FromSubscription(Config.SpeechKey, Config.SpeechRegion);
        config.SpeechRecognitionLanguage = _opts.Language;

        // time-out-uri de liniște (în ms)
        config.SetProperty(PropertyId.SpeechServiceConnection_InitialSilenceTimeoutMs,
                           _opts.InitialSilenceTimeoutMs.ToString());
        config.SetProperty(PropertyId.SpeechServiceConnection_EndSilenceTimeoutMs,
                           _opts.EndSilenceTimeoutMs.ToString());

        recognizer = new SpeechRecognizer(config);
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

        // (re)armează deadline-urile din opțiuni
        StartDeadlines();
    }

    public async Task StopAsync()
    {
        if (!_isListening) return;
        _isListening = false;

        try
        {
            _cts?.Cancel();
            _deadlineCts?.Cancel();
            _deadlineCts = null;

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

    /// <summary>
    /// Actualizează opțiunile în timp ce ascultă.
    /// Aplică live doar deadline-urile; Mode/Limbă/Silence timeouts cer Stop+Start.
    /// </summary>
    public void UpdateOptions(VoiceListenOptions newOpts)
    {
        if (newOpts == null) return;
        var old = _opts;
        _opts = newOpts;

        if (_isListening)
        {
            // live-update: AutoStopAfter / MaxDuration
            bool deadlinesChanged = old.AutoStopAfter != newOpts.AutoStopAfter
                                 || old.MaxDuration != newOpts.MaxDuration;
            if (deadlinesChanged)
                StartDeadlines();

            // restul necesită restart
            if (old.Mode != newOpts.Mode)
                _log.Warn("Mode changed — Stop & Start to apply.");

            if (old.Language != newOpts.Language ||
                old.InitialSilenceTimeoutMs != newOpts.InitialSilenceTimeoutMs ||
                old.EndSilenceTimeoutMs != newOpts.EndSilenceTimeoutMs)
            {
                _log.Warn("Language/Silence timeouts changed — Stop & Start to apply.");
            }
        }
    }

    // ====================== Helpers ======================

    /// <summary>
    /// Armează un singur timer de auto-oprire cu minimul dintre AutoStopAfter și MaxDuration.
    /// </summary>
    private void StartDeadlines()
    {
        if (!_isListening || recognizer == null || _opts.Mode == ListenMode.SingleUtterance)
            return;

        // oprește timerul precedent (dacă există)
        _deadlineCts?.Cancel();

        // niciun deadline setat → nimic de făcut
        if (!_opts.AutoStopAfter.HasValue && !_opts.MaxDuration.HasValue)
            return;

        var deadline = new[]
        {
            _opts.AutoStopAfter ?? TimeSpan.MaxValue,
            _opts.MaxDuration   ?? TimeSpan.MaxValue
        }.Min();

        _deadlineCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token);
        _ = GuardStopAfter(deadline, "[deadline]", _deadlineCts.Token);
    }

    private async Task GuardStopAfter(TimeSpan delay, string tag, CancellationToken token)
    {
        try
        {
            await Task.Delay(delay, token);
            _log.Info($"{tag} elapsed → stopping.");
            await StopAsync();
        }
        catch (TaskCanceledException) { /* rearmat sau oprit manual */ }
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
                await HandleResultAsync(e.Result);
        };

        recognizer.Canceled += (s, e) =>
        {
            if (e.Reason == CancellationReason.Error)
                _log.Error($"Canceled: {e.Reason} | {e.ErrorCode} | {e.ErrorDetails}");
            else
                _log.Warn($"Canceled: {e.Reason}");
        };
    }

    /// <summary>
    /// Procesează textul final: cheamă CLU, routează comanda și o execută.
    /// </summary>
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
