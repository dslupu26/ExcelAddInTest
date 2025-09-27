using ExcelAddInTest.ExcelApi;
using System;

namespace ExcelAddInTest.ExcelApi.Commands
{
    /// <summary>
    /// Bolds all cells in the current selection.
    /// </summary>
    public class BoldSelectedCellsCommand : IExcelCommand
    {
        public void Execute(IExcelActions excel)
        {
            var selection = excel.GetCurrentSelection();
            if (selection != null)
            {
                selection.Font.Bold = true;
            }
        }
    }
}