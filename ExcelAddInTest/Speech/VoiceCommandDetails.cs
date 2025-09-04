using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using static System.Net.Mime.MediaTypeNames;

namespace ExcelAddInTest.Speech
{
    /// <summary>
    ///role of this class is to give a simple, structured way
    ///to pass around the information to commands
    ///raw text from HandleResultAsync() (VoiceInterpreter) ->
    ///-> VoiceCommandParser(parses for intent and cells)
    ///-> VoiceCommandDetails(here !)
    ///which is then a field in every command
     /// </summary>

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
