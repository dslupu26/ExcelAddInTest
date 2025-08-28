using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class SelectAreaCommand
    {
        private readonly string _range;

        public SelectAreaCommand(string range)
        {
            _range = range;
        }
        
        public void Execute(IExcelActions excel)
        {
            excel.SelectRange(_range);
        }
    }
}
