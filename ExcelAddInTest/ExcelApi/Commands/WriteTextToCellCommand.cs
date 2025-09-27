using ExcelAddInTest.ExcelApi;
using System;

namespace ExcelAddInTest.ExcelApi.Commands
{
    /// <summary>
    /// Writes the specified text to the given cell address.
    /// </summary>
    public class WriteTextToCell : IExcelCommand
    {
        private readonly string _address;
        private readonly string _text;

        public WriteTextToCell(string address, string text)
        {
            _address = address ?? throw new ArgumentNullException(nameof(address));
            _text = text ?? throw new ArgumentNullException(nameof(text));
        }

        public void Execute(IExcelActions excel)
        {
            var cell = excel.GetCell(_address);
            if (cell != null)
            {
                cell.Value2 = _text;
            }
        }
    }
}