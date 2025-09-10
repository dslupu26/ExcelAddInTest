using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.ExcelApi
{
    public interface IExcelActions
    {
        void ToggleMainPane();
        void WriteFormula(string address, string formula);

        Excel.Range GetCurrentSelection();
        Excel.Range GetCell(string addr);
        void SelectRange(string fp, string sp);

    }
}
