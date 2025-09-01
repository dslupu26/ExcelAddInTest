using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.ExcelApi.Commands
{
    public class AddSelected
    {
        public void Execute(IExcelActions excel)
        {
            var selection = excel.GetCurrentSelection();

            if (selection == null)
            {
                Console.WriteLine("no selection ??");
                return;
            }

            double sum = 0;


            // for each cell in the selection
            // parse the cell value to a double and output the selection to "val" 

            foreach (Excel.Range cell in selection.Cells)
            {
                try
                {
                    if (double.TryParse(cell.Value?.ToString(), out double val))
                    {
                        sum += val;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"cell reading failed ||| {cell.Address}: {ex.Message}");
                }
            }

            Console.WriteLine($"sum of cells : {sum}");
        }
    }
}
