using ExcelAddInTest.ExcelApi.Commands.Enums;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Documents;
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
        private bool _isDestinationInAddresses = false;
        public AddCells(List<string> addresses, string destination, AddCellsMode mode)
        {
            _addresses = addresses ?? throw new ArgumentNullException(nameof(addresses));
            _destination = destination; // its fine if its null i guess. it just adds and doesnt put it anywhere ig even if its stupid
            _mode = mode;
        }
        /// <summary>
        /// Executes a command to add cells based on the specified mode (present in <see cref="AddCellsMode"/>).
        /// </summary>
        /// <param name="excel"></param>
        public void Execute(IExcelActions excel)
        {
            if (_destination == null)
            {
                _destination = _addresses.LastOrDefault();
                _isDestinationInAddresses = true;
            }
            if (_mode == AddCellsMode.Range)
                AddCellsInRange(excel);
            else if (_mode == AddCellsMode.List)
                AddCellsInList(excel);
            else
                Console.WriteLine("<< add cells command >> unknown type");
        }
        /// <summary>
        /// Adds cells in a rectangular range.
        /// If the destination cell is not in the range, it adds its value to the sum as well.
        /// </summary>
        /// <param name="excel"></param>
        private void AddCellsInRange(IExcelActions excel) 
        {
           
            double sum = 0;
                try
                {
                    var last = _addresses.LastOrDefault();
                    var secondLast = _addresses.Count() > 2 ? _addresses[_addresses.Count() - 2] : null;
                    var end = (!_isDestinationInAddresses && secondLast != null) ? secondLast : last;
                    Excel.Range rs = excel.GetRange(_addresses.First(), end);
                    sum = excel.AddCells(rs);
                    if (excel.IsCellInRange(_destination, _addresses.First(), end) == false)
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
        /// <summary>
        /// Adds cells individually from a list of addresses.
        /// The destination cell is included in the sum.
        /// </summary>
        /// <param name="excel"></param>
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
