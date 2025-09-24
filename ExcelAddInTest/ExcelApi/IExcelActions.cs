using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.ExcelApi
{
    /// <summary>
    /// Used to acces the methods able in the ExcelFacade class.
    /// Available methods that we know how to execute in Excel.
    /// </summary>
    public interface IExcelActions
    {
        void ToggleMainPane();
        void WriteFormula(string address, string formula);
        Excel.Range GetCurrentSelection();
        Excel.Range GetCell(string addr);
        void SelectRange(string fp, string sp);
        void WriteInCell(string cell, string text);
        Excel.Range GetRange(string fp, string sp);
        double AddCells(Excel.Range rs);

        bool IsCellInRange(string cell, string fp, string sp);

    }
}
