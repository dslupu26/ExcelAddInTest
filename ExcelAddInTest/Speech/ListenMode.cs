using ExcelAddInTest;
using System;
public enum ListenMode
{
    SingleUtterance, // se oprește după prima frază (silence detect)
    Continuous       // rămâne pornit până îl oprești tu sau expiră un timer
}

public sealed class VoiceListenOptions
{
    public ListenMode Mode { get; set; } = ListenMode.Continuous;

    // dacă vrei un auto-stop absolut (ex. 15s sau 30s)
    public TimeSpan? AutoStopAfter { get; set; } = null;

    // dacă vrei un „hard cap” când mergi manual (ex. max 30s)
    public TimeSpan? MaxDuration { get; set; } = null;

    // limbă + sensibilitate la pauze
    public string Language { get; set; } = Config.SpeechLanguage; // "en-US"
    public int InitialSilenceTimeoutMs { get; set; } = 5000;  // cât aștepți până la primul sunet
    public int EndSilenceTimeoutMs { get; set; } = 3000;  // câtă liniște înseamnă "sfârșit frază"
}
