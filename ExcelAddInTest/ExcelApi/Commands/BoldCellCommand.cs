using ExcelAddInTest.ExcelApi;
using System;
using System.Collections.Generic;
using System.Windows;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class BoldCellCommand : IExcelCommand
    {
        private readonly List<string> _addresses = new List<string>();

        public BoldCellCommand(List<string> addresses)
        {
            _addresses = addresses ?? throw new ArgumentNullException(nameof(addresses));
        }

        public void Execute(IExcelActions excel)
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
    }
}