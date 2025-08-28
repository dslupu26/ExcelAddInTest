using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.Nlu
{
    public class RangeFromConnector
    {
        public static string BuildRangeAroundConnector(IEnumerable<NluEntity> entities)
        {
            var cells = entities.Where(e => e.Category == "Cell").OrderBy(e => e.Offset).ToList();
            var connectors = entities.Where(e=> e.Category == "RangeConnector").ToList();
            if(cells.Count == 0) return null;

            var conn = connectors.OrderByDescending(c => c.Score).ThenBy(c=>c.Offset).FirstOrDefault();

            NluEntity leftCell = null, rightCell = null;

            if (conn != null) { 
                int connStart = conn.Offset;
                int connEnd = conn.Offset + conn.Length;

                leftCell = cells.Where(c=> c.Offset+ c.Length<= connStart)
                    .OrderByDescending(c => c.Offset + c.Length)
                    .FirstOrDefault();

                rightCell = cells.Where(c => c.Offset >= connEnd)
                    .OrderBy(c => c.Offset)
                    .FirstOrDefault();

            }
            if(leftCell == null || rightCell == null)
            {
                if(cells.Count >= 2)
                {
                    leftCell = cells.First();
                    rightCell = cells.Last();
                }
                else
                {
                    return null;
                }
            }

            if(!TryParseCell(leftCell.Text,out var lc) || !TryParseCell(rightCell.Text, out var rc)){
                return null;
            }
            
            if(lc.col == rc.col)
            {
                var r1 = Math.Min(lc.row, rc.row);
                var r2 = Math.Max(lc.row, rc.row);
                return $"{IndexToCol(lc.col)}{r1}:{IndexToCol(rc.col)}{r2}";

            }
            if (lc.row == rc.row)
            {
                var c1 = Math.Min(lc.col, rc.col);
                var c2 = Math.Max(lc.col, rc.col);
                return $"{IndexToCol(c1)}{lc.row}:{IndexToCol(c2)}{rc.row}";
            }

            var topRow = Math.Min(lc.row, rc.row);
            var bottomRow = Math.Max(lc.row, rc.row);
            var leftCol = Math.Min(lc.col, rc.col);
            var rightCol = Math.Max(lc.col, rc.col);
            return $"{IndexToCol(leftCol)}{topRow}:{IndexToCol(rightCol)}{bottomRow}";

        }

        private static string IndexToCol(int col)
        {
            var sb = new StringBuilder();
            while (col > 0)
            {
                col--;
                sb.Insert(0, (char)('A' + (col % 26)));
                col /= 26;
            }
            return sb.ToString();
        }

        private static bool TryParseCell(string s, out (int col, int row) cell)
        {
            cell = (0, 0);
            if (string.IsNullOrWhiteSpace(s)) return false;

            int i = 0, col = 0;
            while (i < s.Length && char.IsLetter(s[i]))
            {
                col = col * 26 + (char.ToUpperInvariant(s[i]) - 'A' + 1);
                i++;
            }

            int row = 0;
            while (i < s.Length && char.IsDigit(s[i]))
            {
                row = row * 10 + (s[i] - '0');
                i++;
            }

            if (col <= 0 || row <= 0 || i != s.Length) return false;
            cell = (col, row);
            return true;
        }

    }
}
