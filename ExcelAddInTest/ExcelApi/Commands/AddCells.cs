using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelAddInTest.ExcelApi.Commands.Enums;

namespace ExcelAddInTest.ExcelApi.Commands
{
    /// <summary>
    /// Builds a SUM(...) formula and writes it to the destination cell:
    ///  - List mode:  =SUM(A1,A5,B1)
    ///  - Range mode: =SUM(A1:B5)
    /// </summary>
    public sealed class AddCells : IExcelCommand
    {
        private readonly List<string> _addresses;   // normalized A1s (upper-case)
        private readonly string _destination;       // where to write the formula (must be set)
        private readonly AddCellsMode _mode;

        public AddCells(List<string> addresses, string destination, AddCellsMode mode)
        {
            _addresses = addresses ?? throw new ArgumentNullException(nameof(addresses));
            _destination = string.IsNullOrWhiteSpace(destination) ? null : destination.Trim().ToUpperInvariant();
            _mode = mode;
        }

        public void Execute(IExcelActions excel)
        {
            if (excel == null) throw new ArgumentNullException(nameof(excel));
            if (string.IsNullOrWhiteSpace(_destination))
                throw new CommandExecutionException("AddCells needs a destination cell (e.g., '… in B2').");

            switch (_mode)
            {
                case AddCellsMode.Range:
                    WriteRangeFormula(excel);
                    break;

                case AddCellsMode.List:
                default:
                    WriteListFormula(excel);
                    break;
            }
        }

        private void WriteRangeFormula(IExcelActions excel)
        {
            if (_addresses.Count < 2)
                throw new CommandExecutionException("Range mode requires two cells (e.g., 'from A1 to B5').");

            var from = _addresses[0];
            var to = _addresses[1];

            // Avoid circular refs: destination must not be inside the range
            if (excel.IsCellInRange(_destination, from, to))
                throw new CommandExecutionException(
                    $"Destination '{_destination}' cannot be inside the summed range {from}:{to}."
                );

            var formula = $"=SUM({from}:{to})";
            excel.WriteFormula(_destination, formula);
        }

        private void WriteListFormula(IExcelActions excel)
        {
            // Exclude destination from the list to avoid circular references
            var items = _addresses
                .Where(a => !string.Equals(a, _destination, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (items.Count == 0)
                throw new CommandExecutionException("No source cells to sum (list was empty or only contained the destination).");

            // De-duplicate while preserving order (case-insensitive)
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var unique = new List<string>();
            foreach (var a in items)
                if (seen.Add(a)) unique.Add(a);

            var formula = "=SUM(" + string.Join(",", unique) + ")";
            excel.WriteFormula(_destination, formula);
        }
    }
}
