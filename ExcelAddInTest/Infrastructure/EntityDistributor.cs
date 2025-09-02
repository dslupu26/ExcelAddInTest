using Microsoft.CognitiveServices.Speech;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents.DocumentStructures;
using static System.Net.Mime.MediaTypeNames;

namespace ExcelAddInTest
{

    public class EntityDistributor
    {
        private readonly CluService _clu;
        private List<string> cellAddresses;
        private string cellDestination;

        public EntityDistributor(CluService clu)
        {
            _clu = clu;
            cellAddresses = new List<string>();
        }

        public async void ListMaker(SpeechRecognitionResult result)
        {
            var text = result.Text?.Trim();
            bool destinationConnector = false; 

            try
            {
                var nlu = await _clu.AnalyzeAsync(text);

                foreach (var ent in nlu.Entities)
                {
                    if (ent.Category == "Cell")
                        if (destinationConnector == false)
                            cellAddresses.Add(ent.Text);
                        else
                            cellDestination = ent.Text;


                    if (ent.Category == "RangeConnector" && ent.Text.ToLower() == "to")
                        destinationConnector = true;

                    // debugging purposes
                    // Console.WriteLine($" - {ent.Category}: \"{ent.Text}\"");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"something went wrong ||| {ex.Message}");
            }
        }

        // debugging purposes

        public void WriteData()
        {
            MessageBox.Show($"Destination cell : {cellDestination}\n");

            //Console.WriteLine($"Destination cell : {cellDestination}\n");
            //Console.WriteLine($"Cells list : {cellAddresses}");
        }
    }
}
