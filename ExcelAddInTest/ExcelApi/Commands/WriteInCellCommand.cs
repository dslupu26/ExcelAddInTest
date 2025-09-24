using System;
using ExcelAddInTest.ExcelApi;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class WriteInCellCommand : IExcelCommand
    {
        private readonly string _cell;
        private readonly string _text;

        public WriteInCellCommand(string cellA1, string text)
        {
            if (string.IsNullOrWhiteSpace(cellA1))
                throw new ArgumentNullException(nameof(cellA1), "WriteInCell: target cell is required.");

            _cell = cellA1.Trim().ToUpperInvariant();
            _text = text ?? string.Empty;
        }

        public void Execute(IExcelActions excel)
        {
            if (excel == null) throw new ArgumentNullException(nameof(excel));
            excel.WriteInCell(_cell, _text);   // let exceptions bubble to CommandExecutor
        }
    }
}
