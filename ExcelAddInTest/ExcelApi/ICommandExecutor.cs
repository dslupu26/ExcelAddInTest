using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.ExcelApi
{
    public interface ICommandExecutor
    {
        void Execute(IExcelCommand command);
    }
}
