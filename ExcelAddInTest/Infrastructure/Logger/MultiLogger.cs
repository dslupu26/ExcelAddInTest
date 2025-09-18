using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.Infrastructure.Logger
{
    public sealed class MultiLogger : ILogger 
    {
        private readonly IReadOnlyList<ILogger> _targets;
        public MultiLogger(params ILogger[] targets) { _targets = targets; }

        public void Info(string m) { foreach (var t in _targets) t.Info(m); }
        public void Warn(string m) { foreach (var t in _targets) t.Warn(m); }
        public void Error(string m, Exception ex = null) { foreach (var t in _targets) t.Error(m, ex); }
        public void Raw(string m) { foreach (var t in _targets) t.Raw(m); }

        public void Dispose() { foreach (var t in _targets) t.Dispose(); }

    }
}
