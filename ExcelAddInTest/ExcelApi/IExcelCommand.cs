using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.ExcelApi
{
    /// <summary>
    /// Interface used to call commands available in the IExcelActions interface.
    /// </summary>
    public interface IExcelCommand
    {
        void Execute(IExcelActions excel);
    }
}
