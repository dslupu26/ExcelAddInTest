using ExcelAddInTest.ExcelApi.Commands;
using ExcelAddInTest.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;


namespace ExcelAddInTest.ExcelApi
{
    /// <summary>
    /// Executes commands available in the IExcelActions interface through IExcelCommand interface.
    /// </summary>

    public class CommandExecutor : ICommandExecutor
    {
        private IExcelActions _excel;
        private ILogger _log;

        public CommandExecutor(IExcelActions excel, ILogger log)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
            _log = log;
            _log = log;
        }

        public void Execute(IExcelCommand cmd)
        {
            try
            {
                if (cmd == null)
                {
                    _log.Error($"CommandExecutor Error: command is null!");
                    return;
                }
                cmd.Execute(_excel);
            }
            catch (Exception ex)
            {
                _log.Error($"CommandExecutor Error: {ex.Message}");

            }
        }
    }
}
