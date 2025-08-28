using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.ExcelApi
{
    public interface IExcelActions
    {
        void ToggleMainPane();
        void SelectRange(string address);
        void WriteFormula(string address, string formula);
    }
}
