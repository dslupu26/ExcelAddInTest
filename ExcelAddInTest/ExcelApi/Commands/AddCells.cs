using ExcelAddInTest.ExcelApi.Commands.Enums;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class AddCells : IExcelCommand
    {
        private readonly List<string> _addresses;
        private string _destination;
        private EntityDistributor _entityDistrib;
        private AddCellsMode _mode;
        public AddCells(List<string> addresses, string destination, AddCellsMode mode)
        {
            _addresses = addresses ?? throw new ArgumentNullException(nameof(addresses));
            _destination = destination; // its fine if its null i guess. it just adds and doesnt put it anywhere ig even if its stupid
            _mode = mode;
        }

        public void Execute(IExcelActions excel)
        {
            if (_destination == null)
            { 
               _destination = _addresses.LastOrDefault();
            }
            if (_mode == AddCellsMode.Range)
                AddCellsInRange(excel);
            else if (_mode == AddCellsMode.List)
                AddCellsInList(excel);
            else
                Console.WriteLine("<< add cells command >> unknown type");
        }

        private void AddCellsInRange(IExcelActions excel) 
        {
            double sum = 0;
                try
                {
                    Excel.Range rs = excel.GetRange(_addresses.First(), _addresses[_addresses.Count()-2]);
                    sum = excel.AddCells(rs);
                    sum = sum + double.Parse(excel.GetCell(_addresses.Last()).Value2.ToString(), CultureInfo.InvariantCulture);
            }
                catch (Exception ex)
                {
                    Console.WriteLine($"<< add cells command >> cell reading failed |||: {ex.Message}");
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
        private void AddCellsInList(IExcelActions excel)
        {
            double sum = 0;

            foreach (var address in _addresses)
            {
                try
                {
                    double value = 0;
                    Excel.Range cell = excel.GetCell(address);
                    if (double.TryParse(cell.Value2?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out value))
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
