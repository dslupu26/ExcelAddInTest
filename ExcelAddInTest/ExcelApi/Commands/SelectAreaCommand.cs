using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class SelectAreaCommand : IExcelCommand
    {
        private readonly string _fp; // top-left (A1-style)
        private readonly string _sp; // bottom-right (A1-style)

        public SelectAreaCommand(string firstPoint, string secondPoint)
        {
            if (string.IsNullOrWhiteSpace(firstPoint))
                throw new ArgumentNullException(nameof(firstPoint), "<< select area >> first point is null/empty");
            if (string.IsNullOrWhiteSpace(secondPoint))
                throw new ArgumentNullException(nameof(secondPoint), "<< select area >> second point is null/empty");

            // Normalize order to always be TopLeft..BottomRight
            var a = CellRef.Parse(firstPoint);
            var b = CellRef.Parse(secondPoint);
            var tl = CellRef.TopLeft(a, b);
            var br = CellRef.BottomRight(a, b);

            _fp = tl.ToA1();
            _sp = br.ToA1();
        }

        public void Execute(IExcelActions excel)
        {
            // IExcelActions.SelectRange should be implemented as:
            //   var rng = app.Range[firstA1, secondA1]; rng.Select();
            excel.SelectRange(_fp, _sp);
        }

        // ---------------- helpers ----------------

        private struct CellRef
        {
            public int Col; // 1-based (A=1, B=2, …)
            public int Row; // 1-based

            public CellRef(int col, int row) { Col = col; Row = row; }

            private static readonly Regex RxA1 = new Regex(@"^([A-Za-z]{1,3})(\d{1,7})$",
                RegexOptions.Compiled);

            public static CellRef Parse(string a1)
            {
                var m = RxA1.Match(a1.Trim());
                if (!m.Success) throw new ArgumentException($"Invalid A1 address: '{a1}'");
                int col = ColToIndex(m.Groups[1].Value);
                int row = int.Parse(m.Groups[2].Value);
                return new CellRef(col, row);
            }

            public string ToA1() => IndexToCol(Col) + Row.ToString();

            public static CellRef TopLeft(CellRef a, CellRef b) =>
                new CellRef(Math.Min(a.Col, b.Col), Math.Min(a.Row, b.Row));

            public static CellRef BottomRight(CellRef a, CellRef b) =>
                new CellRef(Math.Max(a.Col, b.Col), Math.Max(a.Row, b.Row));

            private static int ColToIndex(string letters)
            {
                int n = 0;
                foreach (char ch in letters.ToUpperInvariant())
                {
                    n = n * 26 + (ch - 'A' + 1);
                }
                return n;
            }

            private static string IndexToCol(int index)
            {
                // 1 -> A, 26 -> Z, 27 -> AA
                var s = "";
                int n = index;
                while (n > 0)
                {
                    n--; // make it 0-based
                    s = (char)('A' + (n % 26)) + s;
                    n /= 26;
                }
                return s;
            }
        }

        // Factory helpers (nice for building from entities)
        public static SelectAreaCommand FromSingle(string cellA1) =>
            new SelectAreaCommand(cellA1, cellA1);

        public static SelectAreaCommand FromTwo(string a1, string b1) =>
            new SelectAreaCommand(a1, b1);

        public static SelectAreaCommand FromCells(IEnumerable<string> cells)
        {
            if (cells == null) throw new ArgumentNullException(nameof(cells));
            var list = cells.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
            if (list.Count == 0) throw new ArgumentException("No cells provided.");
            if (list.Count == 1) return FromSingle(list[0]);

            // Build the bounding rectangle covering ALL cells mentioned
            var parsed = list.Select(CellRef.Parse).ToList();
            int minC = parsed.Min(c => c.Col), maxC = parsed.Max(c => c.Col);
            int minR = parsed.Min(c => c.Row), maxR = parsed.Max(c => c.Row);

            var topLeft = new CellRef(minC, minR).ToA1();
            var botRight = new CellRef(maxC, maxR).ToA1();
            return new SelectAreaCommand(topLeft, botRight);
        }
    }
}
