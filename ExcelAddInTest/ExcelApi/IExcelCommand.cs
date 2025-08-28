using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.ExcelApi
{
    public interface IExcelCommand
    {
        void Execute(IExcelActions excel);
    }
}
