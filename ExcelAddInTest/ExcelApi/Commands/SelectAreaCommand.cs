using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class SelectAreaCommand : IExcelCommand
    {
        private readonly string _fp, _sp;

        public SelectAreaCommand(string firstPoint, string secondPoint)
        {
            _fp = _fp ?? throw new Exception("<< select area command >> first point is null in select area command");
            _sp = _sp ?? throw new Exception("<< select area command >> second point is null in select area command");
        }
        
        public void Execute(IExcelActions excel)
        {
            excel.SelectRange(_fp, _sp);
        }
    }
}
