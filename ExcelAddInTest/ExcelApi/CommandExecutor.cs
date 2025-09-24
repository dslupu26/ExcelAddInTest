using ExcelAddInTest.ExcelApi.Commands;
using ExcelAddInTest.ExcelApi.Commands.Enums;
using ExcelAddInTest.Infrastructure.Logger;
using ExcelAddInTest.Infrastructure.Text;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;


namespace ExcelAddInTest.ExcelApi
{
    /// <summary>
    /// Executes commands available in the IExcelActions interface through IExcelCommand interface.
    /// </summary>

    public static class DictExt
    {

        public static T Get<T>(this Dictionary<string, object> d, string key)
            => d.TryGetValue(key, out var o) && o is T t ? t : default;

        public static bool GetBool(this Dictionary<string, object> d, string key)
           => d.TryGetValue(key, out var o) && o is bool b && b;
    }

    public class CommandExecutor : ICommandExecutor
    {
        private IExcelActions _excel;
        private ILogger _log;
        private IExcelCommand cmd = null;
        public CommandExecutor(IExcelActions excel, ILogger log)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
            _log = log;
        }

        /// <summary>Execute a command; returns true if it completed without throwing.</summary>    
        public bool Execute(IExcelCommand cmd)
        {
            if (cmd == null)
            {
                _log.Error("CommandExecutor: command is null.");
                return false;
            }

            var name = cmd.GetType().Name;
            var sw = Stopwatch.StartNew();
            try
            {
                cmd.Execute(_excel);
                _log.Info($"[{name}] OK in {sw.ElapsedMilliseconds} ms");
                return true;
            }
            catch (CommandExecutionException ex)
            {
                _log.Error($"[{name}] Command execution error.", ex);
            }
            catch (ArgumentException ex)
            {
                _log.Error($"[{name}] Invalid argument(s).", ex);
            }
            catch (COMException ex)
            {
                _log.Error($"[{name}] Excel interop failed.", ex);
            }
            catch (Exception ex)
            {
                _log.Error($"[{name}] Unexpected error.", ex);
            }
            finally
            {
                sw.Stop();
            }

            return false;
        }
    }
}
