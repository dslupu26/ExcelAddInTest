using ExcelAddInTest.ExcelApi;
using ExcelAddInTest.ExcelApi.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace ExcelAddInTest.Nlu
{
    public class IntentRouter
    {
        public IExcelCommand Route(NluModels.NluResult nlu)
        {
            if (nlu == null) return null;

            return null;
        }

    }

}
