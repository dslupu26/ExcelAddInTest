using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class AddCells : IExcelCommand
    {
        private readonly List<string> _addresses;
        private readonly string _destination;

        private EntityDistributor _entityDistrib;

        public AddCells(List<string> addresses, string destination)
        {
            _addresses = addresses ?? throw new ArgumentNullException(nameof(addresses));
            _destination = destination; // its fine if its null i guess. it just adds and doesnt put it anywhere ig even if its stupid
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
                    Console.WriteLine($"<< add cells command >> cell reading failed ||| {address}: {ex.Message}");
                }
            }

            if (string.IsNullOrEmpty(_destination) == false)
            {
                try
                {
                    Excel.Range destinationCell = excel.GetCell(_destination);
                    if (destinationCell != null)
                        destinationCell.Value2 = sum;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"<< add cells command >> could not write to destionation ||| {ex.Message}");
                }
            }
        }
    }
}
