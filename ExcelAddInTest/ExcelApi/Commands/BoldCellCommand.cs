using ExcelAddInTest.ExcelApi;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class BoldCellCommand : IExcelCommand
    {
        private readonly List<string> _addresses = new List<string>();
        private readonly bool _enabled = true;

        public BoldCellCommand(List<string> addresses, bool isRange)
        {
            _addresses = addresses ?? throw new ArgumentNullException(nameof(addresses));
            _enabled = isRange;
        }

        public void Execute(IExcelActions excel)
        {
            if (!_enabled)
            {
                foreach (string address in _addresses)
                {
                    var cell = excel.GetCell(address);
                    if (cell != null)
                    {
                        cell.Font.Bold = true;
                    }
                }
            }
            else if (_enabled && _addresses.Count() == 2)
            {
                var range = excel.GetRange(_addresses.First(), _addresses.Last());
                range.Font.Bold = true;
            }
        }
    }
}
