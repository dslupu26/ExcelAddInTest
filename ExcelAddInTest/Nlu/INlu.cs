using ExcelAddInTest.Nlu.NluModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.Nlu
{
    public interface INlu
    {
        Task<NluResult> AnalyzeAsync(string text);
    }
}
