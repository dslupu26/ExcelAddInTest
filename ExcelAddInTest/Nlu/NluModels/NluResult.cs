using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.Nlu.NluModels
{
    /// <summary>
    /// Represents the result of a natural language understanding operation, holding the 
    /// top intent, a list of the recognized entities and the raw JSON (string) response from the service.
    /// </summary>
    public class NluResult
    {   
        public string TopIntent { get; set; }
        public List<NluEntity> Entities { get; set; }
        public string RawJson { get; set; }
    }
}
