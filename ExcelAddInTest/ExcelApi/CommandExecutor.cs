using ExcelAddInTest.Logging;
using System;

namespace ExcelAddInTest.ExcelApi
{
    /// <summary>
    /// Executes commands available in the IExcelActions interface through IExcelCommand interface.
    /// </summary>
    public class CommandExecutor : ICommandExecutor
    {
        private readonly IExcelActions _excel;
        private readonly ILogger _log;
        public CommandExecutor(IExcelActions excel, ILogger log)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
            _log = log ?? throw new ArgumentNullException(nameof(log));
        }

        /// <summary>
        /// Executes a given command on an Excel instance, logging potential errors if the command fails.
        /// </summary>
        /// <param name="command">The command to be executed on the Excel instance.</param>
        public void Execute(IExcelCommand command)
        { 
            if (command == null) { _log.Error("CommandExecutor - No command routed."); return; }
            try { command.Execute(_excel); }
            catch (Exception ex) { _log.Error("CommandExecutor - Command failed: " + ex); }

        }
    }
}
