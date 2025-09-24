using ExcelAddInTest.ExcelApi;
using ExcelAddInTest.ExcelApi.Commands;
using ExcelAddInTest.ExcelApi.Commands.Enums;
using ExcelAddInTest.Infrastructure.Logger;
using ExcelAddInTest.Infrastructure.Text;
using ExcelAddInTest.Nlu;
using ExcelAddInTest.Nlu.NluModels;
using Microsoft.CognitiveServices.Speech;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Markup;

namespace ExcelAddInTest
{
    public static class DictExt
    {

        public static T Get<T>(this Dictionary<string, object> d, string key)
            => d.TryGetValue(key, out var o) && o is T t ? t : default;

         public static bool GetBool(this Dictionary<string, object> d, string key)
            => d.TryGetValue(key, out var o) && o is bool b && b;
    }

    public class EntityDistributor
    {
        private readonly CluService _clu;
        private readonly object _gate = new object();
        private static readonly Regex CellRx = new Regex(@"\b[A-Z]{1,3}\d{1,7}\b", RegexOptions.Compiled); // Regex idiot for WriteInCell
        // Probably should change it later

        private Dictionary<Type, Dictionary<string, object>> commandEntities;
        private Dictionary<string, Action<IReadOnlyList<NluEntity>>> intentHandler;

        private List<string> cellAddresses;
        private string cellDestination;
        private string intent;
        private string _lastUtterance;

        private ICommandExecutor _executor;
        private ILogger _log;

        public EntityDistributor(CluService clu, ICommandExecutor executor, ILogger log)
        {
            _clu = clu;
            _executor = executor;
            _log = log;
            commandEntities = new Dictionary<Type, Dictionary<string, object>>();

            intentHandler = new Dictionary<string, Action<IReadOnlyList<NluEntity>>>
            {
                ["AddCells"] = HAddCells,
                ["SelectArea"] = HSelectArea,
                ["WriteInCell"] = HWriteInCell,   // <— NEW
                ["Write"] = HWriteInCell,   // (optional alias)
                ["Type"] = HWriteInCell    // (optional alias)
            };

        }

        public async Task AnalyzeAsync(string result)
        {
            var text = result?.Trim();
            if (string.IsNullOrEmpty(text))
                return;

            _lastUtterance = text;

            try
            {
                var nlu = await _clu.AnalyzeAsync(text);
                _log.Raw("[CLU RAW]\r\n" + nlu.RawJson);
                _log.Info("[CLU] TopIntent: " + nlu.TopIntent);

                var ents = nlu.Entities ?? new List<NluEntity>();
                foreach (var ent in ents)
                    _log.Info($" - {ent.Category}: \"{ent.Text}\"");

                var top = nlu.TopIntent ?? "None";

                if (intentHandler.TryGetValue(top, out var handler))
                {
                    try
                    {
                        handler(nlu.Entities ?? new List<NluEntity>());
                    }
                    catch (Exception ex)
                    {
                        _log.Error($"Intent handler '{top}' failed.", ex);
                    }
                }
                else
                {
                    _log.Warn($"No handler registered for intent '{top}'.");
                }
            }
            catch (Exception ex)
            {
                _log.Error("AnalyzeAsync failed.", ex);   // replace Console.WriteLine
            }
        }





        // --------------- Action Handlers ---------------
        //
        //
        // these down here execute the commands (i mean calling the methods not the guillotine)
        // to add more you just. add another handler
        // cheers mate. 
        //
        //
        // we then give this into command executor
        // and just let it do its funky little thing

        private void HAddCells(IReadOnlyList<NluEntity> entities)
        {
            var cell_list = new List<string>();
            var destination = new List<string>();

            bool destinationbool = false;
            bool destinationConnector = false;
            bool rangeConnector = false;
            int rangeConnectorCount = 0;
            bool listConnector = false;

            foreach (var entity in entities)
            {
                switch (entity.Category)
                {
                    case "Cell": cell_list.Add(entity.Text); break;
                    case "RangeConnector": { rangeConnector = true; rangeConnectorCount++; break; }
                    case "ListConnector": listConnector = true; break;
                    case "Destination": { destination.Add(entity.Text); destinationbool = true; break; }
                    case "DestinationConnector": destinationConnector = true; break;
                    case null: break;
                }
            }

            //we need to think about how we want this to operate first so for now the default
            //will be that we add everything to the destination cell, included
            /*if (destinationbool && cell_list.Take(cell_list.Count - 1).Contains(destination.First()))
                cell_list.RemoveAt(cell_list.Count - 1); // remove last cell if its also the destination*/

            commandEntities[typeof(AddCells)] = new Dictionary<string, object>
            {
                ["cells"] = cell_list,
                ["destination"] = destination,
                ["listconnector"] = listConnector,
                ["rangeconnector"] = rangeConnector,
                ["rangeconnectorcount"] = rangeConnectorCount,
                ["destinationconnector"] = destinationConnector
            };
            ExecuteIfPossible(typeof(AddCells));
        }

        private void HWriteInCell(IReadOnlyList<NluEntity> entities)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_lastUtterance))
                    return;

                string cell = entities?
                    .FirstOrDefault(e => string.Equals(e.Category, "Cell", StringComparison.OrdinalIgnoreCase))
                    ?.Text;

                if (string.IsNullOrWhiteSpace(cell))
                {
                    var m = CellRx.Match(_lastUtterance);
                    if (!m.Success) { _log.Warn("WriteInCell: no cell found."); return; }
                    cell = m.Value.ToUpperInvariant();
                }

                var cellOccur = Regex.Match(_lastUtterance, @"\b" + Regex.Escape(cell) + @"\b", RegexOptions.IgnoreCase);
                if (!cellOccur.Success) { _log.Warn("WriteInCell: cell not located in utterance."); return; }

                var tail = _lastUtterance.Substring(cellOccur.Index + cellOccur.Length).Trim();
                tail = Regex.Replace(tail, @"^(in|to|into|with|as|,|:|\-)\s+", "", RegexOptions.IgnoreCase);

                var q = Regex.Match(tail, "^\"([^\"]*)\"|'([^']*)'");
                var textToWrite = q.Success
                    ? (q.Groups[1].Success ? q.Groups[1].Value : q.Groups[2].Value)
                    : tail;

                lock (_gate)
                {
                    commandEntities[typeof(WriteInCellCommand)] = new Dictionary<string, object>
                    {
                        ["cell"] = cell,
                        ["text"] = textToWrite ?? string.Empty
                    };
                }
                ExecuteIfPossible(typeof(WriteInCellCommand));
            }
            catch (Exception ex)
            {
                _log.Error("HWriteInCell failed.", ex);
            }
        }


        private void ExecuteIfPossible(Type t)
        {
            Dictionary<string, object> d;
            lock (_gate)
            {
                if (!commandEntities.TryGetValue(t, out d))
                    return;
            }

            try
            {
                if (t == typeof(AddCells))
                {
                    AddCellsMode mode = AddCellsMode.List;
                    var cells = d.Get<List<string>>("cells") ?? new List<string>();
                    string dest = d.Get<List<string>>("destination")?.FirstOrDefault();
                    var listConnector = d.GetBool("listconnector");
                    var rangeConnector = d.GetBool("rangeconnector");
                    var destConnector = d.GetBool("destinationconnector");
                    var rangeCount = d.Get<int>("rangeconnectorcount");

                    var parsedCells = new List<string>();
                    foreach (var c in cells)
                    {
                        foreach (var parsedCell in TextNormalizer.ExcelCellRegexParser(c))
                            parsedCells.Add(parsedCell);
                    }

                    if (parsedCells.Count == 0)
                    {
                        _log.Warn("[AddCells] No cells recognized.");
                        return;
                    }

                    if (rangeCount > 2 || (rangeConnector && !destConnector))
                        mode = AddCellsMode.Range;

                    _log.Raw($"AddCells will execute in {mode} mode");

                    var cmd = new AddCells(parsedCells, string.IsNullOrEmpty(dest) ? null : dest, mode);
                    _executor.Execute(cmd);   // executor catches runtime errors inside Execute(...)
                }
                else if (t == typeof(SelectAreaCommand))
                {
                    var fp = d.Get<string>("firstPoint");
                    var sp = d.Get<string>("secondPoint");
                    if (string.IsNullOrWhiteSpace(fp) || string.IsNullOrWhiteSpace(sp))
                    {
                        _log.Warn("[SelectArea] Missing endpoints.");
                        return;
                    }

                    try
                    {
                        var cmd = new SelectAreaCommand(fp, sp); // ctor could throw -> catch below
                        _executor.Execute(cmd);
                    }
                    catch (Exception ex)
                    {
                        _log.Error("[SelectArea] Command build failed.", ex);
                    }
                }
                else if (t == typeof(WriteInCellCommand))
                {
                    var cell = d.Get<string>("cell");
                    var text = d.Get<string>("text") ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(cell))
                    {
                        _log.Warn("WriteInCell: missing cell.");
                        return;
                    }

                    try
                    {
                        var cmd = new WriteInCellCommand(cell, text); // ctor could throw -> catch below
                        _executor.Execute(cmd);
                    }
                    catch (Exception ex)
                    {
                        _log.Error("[WriteInCell] Command build failed.", ex);
                    }
                }
            }
            finally
            {
                // prevent reusing stale data if another intent runs later
                lock (_gate) commandEntities.Remove(t);
            }
        }

        private void HSelectArea(IReadOnlyList<NluEntity> entities)
        {
            string firstPoint = null;
            string secondPoint = null;

            var cells = entities.Where(e => e.Category == "Cell").Select(e => e.Text).ToList();


            // this here assumes that we select from A to B, and just that; not multiple areas or other shenanigans

            firstPoint = cells.FirstOrDefault();
            secondPoint = cells.LastOrDefault();


            commandEntities[typeof(SelectAreaCommand)] = new Dictionary<string, object>
            {
                ["firstPoint"] = firstPoint,
                ["secondPoint"] = secondPoint
            };

            ExecuteIfPossible(typeof(SelectAreaCommand));
        }


        // ---------- Get Over Here ----------
        //
        //
        // adding this just to retrieve the dictionaries whenever we call the commands
        // since we're passing them directly to the commands' constructors
        // 
        // this here basically goes
        // "return the correct dictionary. else ah well"

        public T GetEntity<T>(Type commandType, string commandKey)
        {
            if (commandEntities.TryGetValue(commandType, out var dictionary))
            {
                if (dictionary.TryGetValue(commandKey, out var obj) && obj is T t)
                    return t;
            }

            return default;
        }



        // debugging purposes
        // can ignore

        public void WriteData()
        {
            MessageBox.Show($"Destination cell : {cellDestination}\n");

            //Console.WriteLine($"Destination cell : {cellDestination}\n");
            //Console.WriteLine($"Cells list : {cellAddresses}");
        }
    }
}
