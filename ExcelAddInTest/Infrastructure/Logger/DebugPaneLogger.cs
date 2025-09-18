// Logging/DebugPaneLogger.cs
using ExcelAddInTest.Infrastructure.Logger;
using System;

namespace ExcelAddInTest.Infrastructure.Logger
{
    public sealed class DebugPaneLogger : ILogger
    {
        private readonly DebugPane _pane;
        private readonly Func<DateTime> _clock;

        public DebugPaneLogger(DebugPane pane, Func<DateTime> clock = null)
        {
            _pane = pane ?? throw new ArgumentNullException(nameof(pane));
            _clock = clock ?? (() => DateTime.Now);
        }

        private void Write(string line) => _pane.AppendText(line);

        private string Stamp(string level, string msg)
            => $"[{_clock():HH:mm:ss}] {level} {msg}";

        public void Info(string message) => Write(Stamp("[INFO ]", message));
        public void Warn(string message) => Write(Stamp("[WARN ]", message));
        public void Error(string message, Exception ex = null)
            => Write(Stamp("[ERROR]", ex == null ? message : $"{message}\r\n{ex}"));
        public void Raw(string message) => Write(message);

        public void Dispose()
        {
            /* no-op for pane */
        }
    }
}
