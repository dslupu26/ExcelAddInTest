using System;

namespace ExcelAddInTest.ExcelApi.Commands
{
    /// <summary>
    /// Thrown when an Excel command fails to execute (wraps the root cause with context).
    /// </summary>
    public class CommandExecutionException : Exception
    {
        public CommandExecutionException() { }
        public CommandExecutionException(string message) : base(message) { }
        public CommandExecutionException(string message, Exception inner) : base(message, inner) { }
    }
}
