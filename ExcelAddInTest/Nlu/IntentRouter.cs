using ExcelAddInTest.ExcelApi;
using ExcelAddInTest.ExcelApi.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace ExcelAddInTest.Nlu
{
    public class IntentRouter : IIntentRouter
    {

        public IExcelCommand Route(NluModels.NluResult nlu)
        {
            if (nlu == null) return null;

            switch (nlu.TopIntent)
            {
                case "SelectArea":
                    return new SelectAreaCommand(
                       "A1:B5"  // hardcoded for demo purposes
                    );
                default: return null;
            }

        }

    }

}
