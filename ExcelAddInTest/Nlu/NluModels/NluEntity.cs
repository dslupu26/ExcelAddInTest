using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.Nlu
{
    public class NluEntity
    {
        public string Category { get; set; }
        public string Text { get; set; }
        public int Offset { get; set; }
        public int Length { get; set; }
        public float Score { get; set; }
    }
}
