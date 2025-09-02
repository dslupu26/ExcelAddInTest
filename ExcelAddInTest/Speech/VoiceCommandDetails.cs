using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.Speech
{
    // role of this class is to give a simple, structured way
    // to pass around the information to commands
    // raw text from HandleResultAsync() (VoiceInterpreter) ->
    // -> VoiceCommandParser (parses for intent and cells)
    // -> VoiceCommandDetails (here !)
    // which is then a field in every command

    internal class VoiceCommandDetails
    {
        public string rawText;

        public List<string> cells;
        public string topIntent;

        public VoiceCommandDetails(string rT, List<string> addr, string tI)
        {
            rawText = rT;
            cells = addr;
            topIntent = tI;
        }
    }
}
