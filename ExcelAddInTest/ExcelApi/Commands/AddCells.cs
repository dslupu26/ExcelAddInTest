using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class AddCells
    {
        private readonly List<string> _addresses;

        private EntityDistributor _entityDistrib;

        public AddCells(List<string> addresses)
        {
            _addresses = addresses ?? throw new ArgumentNullException(nameof(addresses));
        }

        public void Execute(IExcelActions excel)
        {
            double sum = 0;

            foreach (var address in _addresses)
            {
                try
                {
                    double value = 0;
                    Excel.Range cell = excel.GetCell(address);
                    if (cell != null && double.TryParse(cell.Value2?.ToString(), out value))
                    {
                        sum += value;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"cell reading failed ||| {address}: {ex.Message}");
                }
            }

            Console.WriteLine($"cell sum : {sum}");
        }
    }
}
