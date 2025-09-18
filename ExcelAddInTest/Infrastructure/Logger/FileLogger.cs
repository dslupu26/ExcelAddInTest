using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.Infrastructure.Logger
{
    public sealed class FileLogger : ILogger, IDisposable
    {
        private readonly object _lock = new object();
        private StreamWriter _writer;
        private readonly Func<DateTime> _clock;

        public string Path { get; }

        public FileLogger(string path, Func<DateTime> clock = null)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentNullException(nameof(path));
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path));

            _writer = new StreamWriter(new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            {
                AutoFlush = true
            };

            Path = path;
            _clock = clock ?? (() => DateTime.Now);
            Info($"\n--- Logger started: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} ---");
        }

        private void WriteLine(string line)
        {
            lock (_lock)
            {
                if (_writer != null)
                    _writer.WriteLine(line);
            }
        }

        private string Stamp(string level, string message)
        {
            return $"[{_clock():HH:mm:ss.fff}] {level} {message}";
        }

        public void Info(string msg) => WriteLine(Stamp("[Info]", msg));
        public void Warn(string msg) => WriteLine(Stamp("[Warn]", msg));
        public void Error(string msg, Exception ex = null) => WriteLine(Stamp("[Error]", ex == null ? msg : $"{msg}\r\n{ex}"));
        public void Raw(string msg) => WriteLine(msg);

        public void Dispose()
        {
            lock (_lock)
            {
                if (_writer != null)
                {
                    _writer.Dispose();
                    _writer = null;
                }
            }
        }
    }
}
