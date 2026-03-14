using ExcelAddInTest;
using ExcelAddInTest.ExcelApi;
using ExcelAddInTest.ExcelApi.Commands;

using ExcelAddInTest.Infrastructure.Logger;

using ExcelAddInTest.Nlu;
using ExcelAddInTest.Utils;
using Microsoft.CognitiveServices.Speech;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media.Animation;
using ExcelAddInTest.Infrastructure.Text;

public class VoiceInterpreter
{
    private SpeechRecognizer recognizer;
    private readonly INlu _clu;
    private readonly IExcelActions _exec;
    private readonly ILogger _log;
    private readonly IIntentRouter _intentRouter;

    private CancellationTokenSource _cts;          // current session
    private CancellationTokenSource _deadlineCts;  // auto-stop timer (resettable)
    private bool _isListening;
    private VoiceListenOptions _opts = new VoiceListenOptions();

    private readonly EntityDistributor _ent;
    private Dictionary<string, IExcelCommand> _commands;

    public VoiceInterpreter(
        INlu clu,                           // <- use the interface here
        IExcelActions exec,
        ILogger log,
        EntityDistributor ent,
        IIntentRouter intentRouter)

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
        config.OutputFormat = OutputFormat.Detailed; // require detailed format for NBest list

        // silence time-outs (in ms)
        config.SetProperty(PropertyId.SpeechServiceConnection_InitialSilenceTimeoutMs,
                           _opts.InitialSilenceTimeoutMs.ToString());
        config.SetProperty(PropertyId.SpeechServiceConnection_EndSilenceTimeoutMs,
                           _opts.EndSilenceTimeoutMs.ToString());

       

        if (_opts.Mode == ListenMode.SingleUtterance)
        {
            _log.Info("Started (speak now)");
            try
            {
                var recognizer = new SpeechRecognizer(config);
                var result = await recognizer.RecognizeOnceAsync().WithCancellation(_cts.Token);
                var bestText = NormalizeForExcel(result);
                await HandleResultAsync(bestText);
                recognizer?.Dispose();
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

        recognizer = new SpeechRecognizer(config);
        WireEvents();
        // Continuous
        await recognizer.StartContinuousRecognitionAsync();
        _log.Info("Started (speak now)");

        // (re)arm deadlines from options
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
    /// Updates the options while listening.
    /// Only applies deadlines live; Mode/Language/Silence timeouts require Stop+Start.
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

            // the rest require a restart
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
    /// Arms a single auto-stop timer with the minimum of AutoStopAfter and MaxDuration.
    /// </summary>
    private void StartDeadlines()
    {
        if (!_isListening || recognizer == null || _opts.Mode == ListenMode.SingleUtterance)
            return;

        // stop the previous timer (if it exists)
        _deadlineCts?.Cancel();

        // no deadline set → nothing to do
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
        catch (TaskCanceledException) { /* rearmed or stopped manually */ }
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
                var bestText = NormalizeForExcel(e.Result);
                await HandleResultAsync(bestText);
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

    
    
    private string NormalizeForExcel(SpeechRecognitionResult e)
    {
        if (string.IsNullOrEmpty(e.Text))
        {
            _log.Warn("No speech recognized.");
            return "";
        }
        // detailed results are in the JSON
        var json = e.Properties.GetProperty(PropertyId.SpeechServiceResponse_JsonResult);

        // extract NBest list containing alternative texts and confidence scores
        JsonDocument doc = JsonDocument.Parse(json);
        var nbest = doc.RootElement.GetProperty("NBest").EnumerateArray()
                                   .Select(selector => new
                                   {
                                       Text = selector.GetProperty("Display").GetString(),
                                       Confidence = selector.TryGetProperty("Confidence", out var conf) ? conf.GetDouble() : 0.0
                                   }).ToList();
        _log.Info("NBest alternatives:");
        foreach (var alt in nbest)
            _log.Info($" - \"{alt.Text}\" (Confidence: {alt.Confidence})");

        //Regex for cell addresses (e.g. A1, B2, AA10, etc.)
        var cellRx = new Regex(@"\b[A-Z]{1,3}[0-9]{1,3}\b");

        // pick the alternative that has the most cell addresses, then by confidence
        var pick = nbest
            .OrderByDescending(x => cellRx.Matches(x.Text ?? "").Count)
            .ThenByDescending(x => x.Confidence)
            .FirstOrDefault();

        var bestText = TextNormalizer.Normalize(pick?.Text ?? e.Text);  // apply lexicon/pipeline here
        return bestText;
    }

    /// <summary>
    /// Processes the final text: calls CLU, routes the command and executes it.
    /// </summary>
    private async Task HandleResultAsync(string result)
    {
        var text = result.Trim(); 
        // this line here has the FINAL result
        _log.Info("Final: " + text);
        // ignore empty results to avoid unnecessary CLU calls
        try
        {
            if (!string.IsNullOrWhiteSpace(text))
            {
                /*var nlu = await _clu.AnalyzeAsync(text);

                _log.Raw("[CLU RAW]\r\n" + nlu.RawJson);
                _log.Info("[CLU] TopIntent: " + nlu.TopIntent);
                // nlu.Entities has the cells; lowkey no need to parse them. again.
                foreach (var ent in nlu.Entities)
                    _log.Info($" - {ent.Category}: \"{ent.Text}\"");


                string intent = nlu.TopIntent;*/

                await _ent.AnalyzeAsync(result);
            }
        }
        catch (Exception exClu)
        {
            _log.Error("[CLU] ERROR", exClu);
        }

    }
}
