using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelAddInTest.ExcelApi;

namespace ExcelAddInTest.ExcelApi.Commands
{
    /// <summary>
    /// Numerically adds SUM(sources) into destination cell (in-place).
    /// No formula is written (avoids circular references).
    /// </summary>
    public sealed class AddIntoCellCommand : IExcelCommand
    {
        private readonly List<string> _sources;
        private readonly string _dest;

        public AddIntoCellCommand(IEnumerable<string> sources, string destination)
        {
            if (string.IsNullOrWhiteSpace(destination))
                throw new ArgumentNullException(nameof(destination));

            _dest = destination.Trim().ToUpperInvariant();
            _sources = new List<string>();
            if (sources != null)
            {
                foreach (var s in sources)
                {
                    if (!string.IsNullOrWhiteSpace(s))
                        _sources.Add(s.Trim().ToUpperInvariant());
                }
            }

            if (_sources.Count == 0)
                throw new ArgumentException("AddInto needs at least one source cell.");
        }

        public void Execute(IExcelActions excel)
        {
            if (excel == null) throw new ArgumentNullException(nameof(excel));

            // sum sources
            double add = 0;
            foreach (var s in _sources)
            {
                var v = excel.GetCell(s)?.Value2;

                double parsed;
                if (TryToDouble(v, out parsed))
                    add += parsed;
            }

            // read current dest
            var curObj = excel.GetCell(_dest)?.Value2;

            double curParsed;
            var cur = TryToDouble(curObj, out curParsed) ? curParsed : 0;

            var total = cur + add;

            // write back as value (not formula)
            excel.WriteInCell(_dest, total.ToString(CultureInfo.InvariantCulture));
        }

        private static bool TryToDouble(object value, out double d)
        {
            d = 0;
            if (value == null) return false;

            try
            {
                // common Excel Value2 cases
                if (value is double) { d = (double)value; return true; }
                if (value is float) { d = (float)value; return true; }
                if (value is int) { d = (int)value; return true; }
                if (value is long) { d = (long)value; return true; }
                if (value is decimal) { d = (double)(decimal)value; return true; }

                var s = value as string;
                if (s != null)
                    return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d);

                d = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
