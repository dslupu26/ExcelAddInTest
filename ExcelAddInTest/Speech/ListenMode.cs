using ExcelAddInTest;
using System;
public enum ListenMode
{
    SingleUtterance, // stops after the first phrase (silence detect)
    Continuous       // stays on until you stop it or a timer expires
}

public sealed class VoiceListenOptions
{
    public ListenMode Mode { get; set; } = ListenMode.Continuous;

    // if you want an absolute auto-stop (e.g. 15s or 30s)
    public TimeSpan? AutoStopAfter { get; set; } = null;

    // if you want a "hard cap" when going manual (e.g. max 30s)
    public TimeSpan? MaxDuration { get; set; } = TimeSpan.FromSeconds(30);

    // language + sensitivity to pauses
    public string Language { get; set; } = Config.SpeechLanguage; // "en-US"
    public int InitialSilenceTimeoutMs { get; set; } = 5000;  // how long to wait for the first sound
    public int EndSilenceTimeoutMs { get; set; } = 3000;  // how much silence means "end of phrase"
}
