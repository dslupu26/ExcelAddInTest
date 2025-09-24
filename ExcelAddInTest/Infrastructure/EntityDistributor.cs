using ExcelAddInTest.ExcelApi;
using ExcelAddInTest.ExcelApi.Commands;
using ExcelAddInTest.ExcelApi.Commands.Enums;
using ExcelAddInTest.Infrastructure.Logger;
using ExcelAddInTest.Infrastructure.Text;
using ExcelAddInTest.Nlu;
using ExcelAddInTest.Nlu.NluModels;
using Microsoft.CognitiveServices.Speech;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Markup;

namespace ExcelAddInTest
{
  

    public class EntityDistributor
    {
        private readonly CluService _clu;

        private Dictionary<string, object> commandData;
        private Dictionary<string, Action<IReadOnlyList<NluEntity>>> intentHandler;

        private List<string> cellAddresses;
        private string cellDestination;
        private string intent;

        private ICommandExecutor _executor;
        private ILogger _log;

        public EntityDistributor(CluService clu, ICommandExecutor executor, ILogger log)
        {
            _clu = clu;
            _executor = executor;
            _log = log;
            commandData = new Dictionary<string, object>();

            intentHandler = new Dictionary<string, Action<IReadOnlyList<NluEntity>>>
            {
                ["AddCells"] = HAddCells,
                ["SelectArea"] = HSelectArea
            };

        }

        public async Task AnalyzeAsync(string result)
        {
            var text = result.Trim();
            if (string.IsNullOrEmpty(text))
                return;

            try
            {
                text = "Add A1 and B7 in D1"; // debug line
                var nlu = await _clu.AnalyzeAsync(text);
                _log.Raw("[CLU RAW]\r\n" + nlu.RawJson);
                _log.Info("[CLU] TopIntent: " + nlu.TopIntent);
                // nlu.Entities has the cells; lowkey no need to parse them. again.
                foreach (var ent in nlu.Entities)
                    _log.Info($" - {ent.Category}: \"{ent.Text}\"");


                string intent = nlu.TopIntent ?? "None";

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
            var cell_list = new List<string>();
            var destination = new List<string>();

            bool destinationbool = false;
            bool destinationConnector = false;
            bool rangeConnector = false;
            int rangeConnectorCount = 0;
            bool listConnector = false;

            foreach (var entity in entities)
            {
                switch (entity.Category)
                {
                    case "Cell": cell_list.Add(entity.Text); break;
                    case "RangeConnector": { rangeConnector = true; rangeConnectorCount++; break; }
                    case "ListConnector": listConnector = true; break;
                    case "Destination": { destination.Add(entity.Text); destinationbool = true; break; }
                    case "DestinationConnector": destinationConnector = true; break;
                    case null: break;
                }
            }

            //we need to think about how we want this to operate first so for now the default
            //will be that we add everything to the destination cell, included
            /*if (destinationbool && cell_list.Take(cell_list.Count - 1).Contains(destination.First()))
                cell_list.RemoveAt(cell_list.Count - 1); // remove last cell if its also the destination*/

            commandData = new Dictionary<string, object>
            {
                ["cells"] = cell_list,
                ["destination"] = destination,
                ["listconnector"] = listConnector,
                ["rangeconnector"] = rangeConnector,
                ["rangeconnectorcount"] = rangeConnectorCount,
                ["destinationconnector"] = destinationConnector
            };

            _executor.Execute(typeof(AddCells), commandData);
        }

        private void HSelectArea(IReadOnlyList<NluEntity> entities)
        {
            string firstPoint = null;
            string secondPoint = null;

            var cells = entities.Where(e => e.Category == "Cell").Select(e => e.Text).ToList();


            // this here assumes that we select from A to B, and just that; not multiple areas or other shenanigans

            firstPoint = cells.FirstOrDefault();
            secondPoint = cells.LastOrDefault();


            commandData = new Dictionary<string, object>
            {
                ["firstPoint"] = firstPoint,
                ["secondPoint"] = secondPoint
            };
            _executor.Execute(typeof(SelectAreaCommand), commandData);
        }


        // ---------- Get Over Here ----------
        //
        //
        // adding this just to retrieve the dictionaries whenever we call the commands
        // since we're passing them directly to the commands' constructors
        // 
        // this here basically goes
        // "return the correct dictionary. else ah well"


        //need to review this later
        /* public T GetEntity<T>(Type commandType, string commandKey)
         {
             if (commandData.TryGetValue(commandType, out var dictionary))
             {
                 if (dictionary.TryGetValue(commandKey, out var obj) && obj is T t)
                     return t;
             }

             return default;
         }*/



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
