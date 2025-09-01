using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.Nlu
{
    /// <summary>
    /// Represents an entity recognized in natural language understanding, 
    /// including its category, text, position, length, and confidence score.
    /// </summary>
    public class NluEntity
    {
        public string Category { get; set; }
        public string Text { get; set; }
        public int Offset { get; set; }
        public int Length { get; set; }
        public float Score { get; set; }
    }
}
