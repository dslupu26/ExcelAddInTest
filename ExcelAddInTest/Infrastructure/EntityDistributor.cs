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


    public class EntityDistributor
    {
        private readonly CluService _clu;
        private readonly object _gate = new object();
        private static readonly Regex CellRx = new Regex(@"\b[A-Z]{1,3}\d{1,7}\b", RegexOptions.Compiled); // Regex idiot for WriteInCell
        // Probably should change it later

        private Dictionary<string, object> commandData;
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
            commandData = new Dictionary<string, object>();

            intentHandler = new Dictionary<string, Action<IReadOnlyList<NluEntity>>>
            {
                ["AddCells"] = HAddCells,
                ["SelectArea"] = HSelectArea,
                ["WriteInCell"] = HWriteInCell,   // <— NEW
                ["WriteValue"] = HWriteInCell,   // (optional alias)
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
            var cell_list_parsed = new List<string>();
            var destination = new List<string>();

            bool rangeConnector = false;
            int rangeConnectorCount = 0;
            bool listConnector = false;

            foreach (var entity in entities)
            {
                switch (entity.Category)
                {
                    case "Cell":
                        cell_list.Add(entity.Text);
                        break;

                    case "RangeConnector":
                        rangeConnector = true;
                        rangeConnectorCount++;
                        break;

                    case "ListConnector":
                        listConnector = true;
                        break;

                }
            }

            foreach (var cell in cell_list)
            {
                foreach (var res_cell in TextNormalizer.ExcelCellRegexParser(cell))
                    cell_list_parsed.Add(res_cell);
            }
            // Heuristics based on the actual words the user spoke:
            var utter = _lastUtterance ?? string.Empty;
            var hasInWord = Regex.IsMatch(utter, @"\b(in|into|at)\b", RegexOptions.IgnoreCase);
            var hasToWord = Regex.IsMatch(utter, @"\bto\b", RegexOptions.IgnoreCase);
            var ToWordCount = Regex.Matches(utter, @"\bto\b", RegexOptions.IgnoreCase).Count;


            // ---- Case 1: "Add A1 to B1" (no "in/into") => in-place add (no formula) ----
            if (!hasInWord && ToWordCount == 1 && cell_list_parsed.Count == 2 && !listConnector)
            {
                var dest = cell_list_parsed.LastOrDefault();
                var sources = cell_list_parsed.GetRange(0, cell_list.Count - 1);

                lock (_gate)
                {
                    commandData = new Dictionary<string, object>
                    {
                        ["sources"] = sources,
                        ["dest"] = dest
                    };
                }
                _executor.Execute(typeof(AddIntoCellCommand), commandData);
                return;
            }

            // ---- Case 2: Formula writer ("... in C1") ----
            // If CLU didn't tag Destination but we DO have an "in/into", infer last cell as destination.
            if (cell_list_parsed.Count > 2)
            {
                var inferred = cell_list_parsed.Last();
                destination.Add(inferred);
                _log.Info($"[AddCells] Inferred destination '{inferred}' after 'in/into'.");
                // optional: remove from sources
                // cell_list.RemoveAt(cell_list.Count - 1);
            }

            lock (_gate)
            {
                commandData = new Dictionary<string, object>
                {
                    ["cells"] = cell_list_parsed,
                    ["destination"] = destination,
                    ["listconnector"] = listConnector,
                    ["rangeconnector"] = rangeConnector,
                    ["rangeconnectorcount"] = rangeConnectorCount,
                    ["towordcount"] = ToWordCount
                };
            }

            _executor.Execute(typeof(AddCells), commandData);
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
                    commandData = new Dictionary<string, object>
                    {
                        ["cell"] = cell,
                        ["text"] = textToWrite ?? string.Empty
                    };
                }
                _executor.Execute(typeof(WriteInCellCommand), commandData);
            }
            catch (Exception ex)
            {
                _log.Error("HWriteInCell failed.", ex);
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


            commandData = new Dictionary<string, object>
            {
                ["firstPoint"] = firstPoint,
                ["secondPoint"] = secondPoint
            };
            _executor.Execute(typeof(SelectAreaCommand), commandData);
        }


        // ---------- Get Over Here ----------
        //
        //
        // adding this just to retrieve the dictionaries whenever we call the commands
        // since we're passing them directly to the commands' constructors
        // 
        // this here basically goes
        // "return the correct dictionary. else ah well"


        //need to review this later
        /* public T GetEntity<T>(Type commandType, string commandKey)
         {
             if (commandData.TryGetValue(commandType, out var dictionary))
             {
                 if (dictionary.TryGetValue(commandKey, out var obj) && obj is T t)
                     return t;
             }

             return default;
         }*/



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
