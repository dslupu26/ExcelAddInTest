﻿using ExcelAddInTest.ExcelApi.Commands;
using ExcelAddInTest.ExcelApi.Commands.Enums;
using ExcelAddInTest.Infrastructure.Logger;
using System;
using System.Collections.Generic;
using System.Linq;


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
        private readonly object _gate = new object();
        public CommandExecutor(IExcelActions excel, ILogger log)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
            _log = log;
        }

        /// <summary>Execute a command; returns true if it completed without throwing.</summary>    
        public bool Execute(Type t, Dictionary<string,object> d)
        {
            if (t == null)
            {
                _log.Error($"CommandExecutor Error: command type is null!");
                return false;
            }

            try
            {
                if (t == typeof(AddCells))
                {
                    var cells = d.Get<List<string>>("cells") ?? new List<string>();
                    var dest = cells.LastOrDefault();
                    var hasRangeConnector = d.GetBool("rangeconnector");
                    var toWordCount = d.Get<int>("towordcount");
                    var listConnector = d.Get<bool>("listconnector");

                    AddCellsMode mode;
                    // The case with 1 "to" and 2 cells is handled in EntityDistributor => AddIntoCellsCommand
                    // Rules:
                    // 1. "to" used >= 2 times with > 2 cells => range
                    // 2. "to" used once, 3 cells, no list connector => range
                    // 3. "to" used once, with list connector => list
                    // 4. all other cases => list (including cell enumerations)
                    if (toWordCount >= 2 && cells.Count() > 2) {mode = AddCellsMode.Range;}
                    else if (toWordCount == 1 && cells.Count == 3 && !listConnector) {mode = AddCellsMode.Range;}
                    else if(toWordCount == 1 && listConnector) { mode = AddCellsMode.List; }
                    else { mode = AddCellsMode.List; }

                  /*// Range if we have a range connector and at least two cells; otherwise List
                    mode = (hasRangeConnector && parsedCells.Count >= 2)
                                ? AddCellsMode.Range
                                : AddCellsMode.List;*/

                    // Destination required (per your spec "... in B2/C1")
                    if (string.IsNullOrWhiteSpace(dest))
                    {
                        _log.Warn("[AddCells] Missing destination cell (say for example: '… in B2').");
                        return false;
                    }

                    var cmd = new AddCells(cells, dest, mode);
                    cmd.Execute(_excel);
                    return true;
                }
                else if (t == typeof(SelectAreaCommand))
                {
                    var fp = d.Get<string>("firstPoint");
                    var sp = d.Get<string>("secondPoint");
                    if (string.IsNullOrWhiteSpace(fp) || string.IsNullOrWhiteSpace(sp))
                    {
                        _log.Warn("[SelectArea] Missing endpoints.");
                        return false;
                    }

                    try
                    {
                        var cmd = new SelectAreaCommand(fp, sp);
                        cmd.Execute(_excel);
                    }
                    catch (Exception ex)
                    {
                        _log.Error("[SelectArea] Command build failed.", ex);
                        return false;
                    }
                    return true;
                }
                else if (t == typeof(AddIntoCellCommand))
                {
                    var sources = d.Get<List<string>>("sources") ?? new List<string>();
                    var dest = d.Get<string>("dest");

                    if (string.IsNullOrWhiteSpace(dest) || sources.Count == 0)
                    {
                        _log.Warn("[AddInto] Missing destination or sources.");
                        return false;
                    }

                    try
                    {
                        var cmd = new AddIntoCellCommand(sources, dest);
                        cmd.Execute(_excel);
                    }
                    catch (Exception ex)
                    {
                        _log.Error("[AddInto] Command build failed.", ex);
                    }
                    return true;
                }
                else if (t == typeof(WriteInCellCommand))
                {
                    var cell = d.Get<string>("cell");
                    var text = d.Get<string>("text") ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(cell))
                    {
                        _log.Warn("WriteInCell: missing cell.");
                        return false;
                    }

                    try
                    {
                        var cmd = new WriteInCellCommand(cell, text);
                        cmd.Execute(_excel);
                    }
                    catch (Exception ex)
                    {
                        _log.Error("[WriteInCell] Command build failed.", ex);
                    }
                    return true;
                }
                else if(t == typeof(BoldCellCommand))
                {
                    var cells = d.Get<List<string>>("Cells");
                    var range = d.Get<bool>("RangeConnector");

                    try
                    {
                        if (range == true && cells.Count != 2)
                            throw new Exception(@"CommandExecutor >> do NOT say something along the lines of ""bold from A1 to B5C6""");

                        var cmd = new BoldCellCommand(cells, range);
                        cmd.Execute(_excel);
                    }
                    catch (Exception ex)
                    {
                        _log.Error("CommandExecutor >> [BoldCells] no work pls call Handy Manny ||| " + ex.Message);
                    }
                    return true;

                }
            }
            finally
            {
                // leaving this here just in case
            }

            return false;
        }
    }
}
