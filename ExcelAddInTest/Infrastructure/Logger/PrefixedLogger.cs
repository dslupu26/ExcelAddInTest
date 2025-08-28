// Logging/PrefixedLogger.cs
using System;

namespace ExcelAddInTest.Logging
{
    public sealed class PrefixedLogger : ILogger
    {
        private readonly ILogger _inner;
        private readonly string _prefix; // e.g., "[Speech] " or "[CLU] "

        public PrefixedLogger(ILogger inner, string prefix)
        {
            _inner = inner;
            _prefix = string.IsNullOrEmpty(prefix) ? "" : prefix + " ";
        }

        public void Info(string m) => _inner.Info(_prefix + m);
        public void Warn(string m) => _inner.Warn(_prefix + m);
        public void Error(string m, Exception ex = null) => _inner.Error(_prefix + m, ex);
        public void Raw(string m) => _inner.Raw(m); // RAW stays raw (often JSON)
    }
}
