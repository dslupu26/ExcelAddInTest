using ExcelAddInTest.ExcelApi.Commands;
using ExcelAddInTest.Nlu;
using Microsoft.CognitiveServices.Speech;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelAddInTest
{

    public class EntityDistributor
    {
        private readonly CluService _clu;

        private Dictionary<Type, Dictionary<string, object>> commandEntities;
        private Dictionary<string, Action<IReadOnlyList<NluEntity>>> intentHandler;


        private List<string> cellAddresses;
        private string cellDestination;
        private string intent;

        public EntityDistributor(CluService clu)
        {
            _clu = clu;

            commandEntities = new Dictionary<Type, Dictionary<string, object>>();

            intentHandler = new Dictionary<string, Action<IReadOnlyList<NluEntity>>>
            {
                ["AddCells"] = HAddCells,
                ["SelectArea"] = HSelectArea
            };
        }

        public async Task AnalyzeAsync(SpeechRecognitionResult result)
        {
            var text = result.Text?.Trim();
            if (string.IsNullOrEmpty(text))
                return;

            try
            {
                var nlu = await _clu.AnalyzeAsync(text);

                string intent = null;

                foreach (var ent in nlu.Entities)
                    if (ent.Category.ToLower() == "topintent")
                        intent = ent.Text;

                if (intentHandler.TryGetValue(intent, out var handler))
                    handler(nlu.Entities);

            }
            catch (Exception ex)
            {
                Console.WriteLine("something important exploded - entity distributor");
            }
        }





        // --------------- Action Handlers ---------------
        //
        //
        // these down here execute the commands (i mean calling the methods not the guillotine)
        // to add more you just. add another handler
        // cheers mate. 
        //
        //
        // we then give this into command executor
        // and just let it do its funky little thing

        private void HAddCells(IReadOnlyList<NluEntity> entities)
        {
            var factors = new List<string>();

            string destination = null;

            bool destinationConnector = false;

            foreach (var entity in entities)
            {
                if (entity.Category == "Cell")
                {
                    if (!destinationConnector)
                        factors.Add(entity.Text);
                    else
                        destination = entity.Text;
                }

                if (entity.Category == "RangeConnector" && entity.Text.Equals("to"))
                    destinationConnector = true;
            }

            commandEntities[typeof(AddCells)] = new Dictionary<string, object>
            {
                ["factors"] = factors,
                ["destination"] = destination
            };
        }

        private void HSelectArea(IReadOnlyList<NluEntity> entities)
        {
            string firstPoint = null;
            string secondPoint = null;

            var cells = entities.Where(e => e.Category == "Cell").Select(e => e.Text).ToList();


            // this here assumes that we select from A to B, and just that; not multiple areas or other shenanigans

            if (cells.Count() == 2)
            {
                firstPoint = cells[0];
                secondPoint = cells[1];
            }

            commandEntities[typeof(SelectAreaCommand)] = new Dictionary<string, object>
            {
                ["firstPoint"] = firstPoint,
                ["secondPoint"] = secondPoint
            };
        }


        // ---------- Get Over Here ----------
        //
        //
        // adding this just to retrieve the dictionaries whenever we call the commands
        // since we're passing them directly to the commands' constructors
        // 
        // this here basically goes
        // "return the correct dictionary. else ah well"

        public T GetEntity<T>(Type commandType, string commandKey)
        {
            if (commandEntities.TryGetValue(commandType, out var dictionary))
            {
                if (dictionary.TryGetValue(commandKey, out var obj) && obj is T t)
                    return t;
            }

            return default;
        }



        // debugging purposes
        // can ignore

        public void WriteData()
        {
            MessageBox.Show($"Destination cell : {cellDestination}\n");

            //Console.WriteLine($"Destination cell : {cellDestination}\n");
            //Console.WriteLine($"Cells list : {cellAddresses}");
        }
    }
}
