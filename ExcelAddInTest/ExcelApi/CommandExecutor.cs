using ExcelAddInTest.ExcelApi.Commands;
using ExcelAddInTest.ExcelApi.Commands.Enums;
using ExcelAddInTest.Infrastructure.Logger;
using ExcelAddInTest.Infrastructure.Text;
using System;
using System.CodeDom.Compiler;
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

        public void Execute(Type t, Dictionary<string,object> d)
        {
            if (t == null)
            {
                _log.Error($"CommandExecutor Error: command type is null!");
                return;
            }

            

            if (t == typeof(AddCells))
            {
                AddCellsMode mode = AddCellsMode.List;
                var cells = d.Get<List<string>>("cells");
                string dest = d.Get<List<string>>("destination").FirstOrDefault();
                var listConnector = d.GetBool("listconnector");
                var rangeConnector = d.GetBool("rangeconnector");
                var destConnector = d.GetBool("destinationconnector");
                var rangeConnectorCount = d.Get<int>("rangeconnectorcount");

                var parsedCells = new List<string>();
                foreach (var c in cells)
                {
                    foreach (var parsedCell in TextNormalizer.ExcelCellRegexParser(c))
                        parsedCells.Add(parsedCell);
                }

                //redundant dar eh...
                if (listConnector is true)
                    mode = AddCellsMode.List;
                else if (rangeConnectorCount > 2 || rangeConnector && destConnector)
                    mode = AddCellsMode.Range;

                _log.Raw($"AddCells command will execute the {mode} version");
                if (string.IsNullOrEmpty(dest))
                    dest = null;
                cmd = new AddCells(parsedCells, dest, mode);
                     
            }

            if (t == typeof(SelectAreaCommand))
            {
                var fp = d.Get<string>("firstPoint");
                var sp = d.Get<string>("secondPoint");

                cmd = new SelectAreaCommand(fp, sp);
            }

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
