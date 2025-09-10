using ExcelAddInTest.ExcelApi.Commands;
using ExcelAddInTest.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelAddInTest.ExcelApi
{
    /// <summary>
    /// Executes commands available in the IExcelActions interface through IExcelCommand interface.
    /// </summary>
    public static class CommandExecutor
    {

        private static readonly Dictionary<string, Func<EntityDistributor, IExcelCommand>> commandFactory =
            new Dictionary<string, Func<EntityDistributor, IExcelCommand>>
            {
                {
                    "AddCells", d => new AddCells(
                        d.GetEntity<List<string>>(typeof(AddCells), "factors"),
                        d.GetEntity<string>(typeof(AddCells), "destination"))
                },

                {
                    "SelectAreaCommand", d => new SelectAreaCommand(
                        d.GetEntity<string>(typeof(SelectAreaCommand), "firstPoint"),
                        d.GetEntity<string>(typeof(SelectAreaCommand), "secondPoint"))
                }
            };




        /// <summary>
        /// Executes a given command on an Excel instance, logging potential errors if the command fails.
        /// </summary>
        /// <param name="command"></param>
        /// <param name="_log"></param>
        /// <param name="_excel"></param>
 
        public static void ExecuteIntent(string intent, EntityDistributor distributor, ILogger _log, IExcelActions _excel)
        { 
            if (intent == null)
            {
                _log.Error("<< command executor >> intent equal to null. what ?");
                return;
            }

            if (distributor == null)
            {
                _log.Error("<< command executor >> distributor equal to null. are you sure you've called  ` = new EntityDistributor(_clu); ` somewhere ?");
                return;
            }

            if (_excel == null)
            {
                _log.Error("<< command executor >> challening my inner arch user, i tell you : skill issue. how ? (provide _excel properly in the constructor)");
                return;
            }

            if (_log == null)
            {
                MessageBox.Show("<< command executor >> this is hilarious to me. how did you even manage to pull this off ?? (_log = null)");
            }

            try
            {
                if (!commandFactory.TryGetValue(intent, out var factory))
                {
                    _log.Error("<< command executor >> intent broken pls fix" + intent);
                    return;
                }

                var command = factory(distributor);

                command.Execute(_excel);
            }
            catch (Exception ex)
            {
                _log.Error("<< command executor >> command failed, see error : " + ex.Message);
            }



        }
    }
}
