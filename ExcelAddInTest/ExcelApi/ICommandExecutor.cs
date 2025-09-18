using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.ExcelApi
{
    /// <summary>
    /// Used to acces the methods able in the CommandExecutor class.
    /// </summary>
    public interface ICommandExecutor
    {
        void Execute(IExcelCommand command);

    }
}
