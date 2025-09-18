using Microsoft.CognitiveServices.Speech.Dialog;
using System;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.ExcelApi
{
    /// <summary>
    /// Interface used for interacting with Excel application and ensuring thread safety.
    /// </summary>
    public sealed class ExcelFacade : IExcelActions
    {
        private readonly Excel.Application _app;
        private Excel.Worksheet ws;
        private readonly Microsoft.Office.Tools.CustomTaskPane _pane;
        private readonly SynchronizationContext _ctx;

        public ExcelFacade(Excel.Application app, Microsoft.Office.Tools.CustomTaskPane pane, SynchronizationContext ctx)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _pane = pane ?? throw new ArgumentNullException(nameof(pane));
            _ctx = ctx ?? SynchronizationContext.Current;
            ws = _app.ActiveSheet;
        }

        public void SetSheet()
        {
            ws = _app.ActiveSheet;
        }
        // this returns nothing
        // therefore it's only good when you don't want to get back *something*
        // kept it here because I don't know if the code explodes without it

        private void OnUi(Action action)
        {
            if (SynchronizationContext.Current == _ctx) {
                action();
                return;
            }

            // this runs action() on the UI thread
            // ignores whatever object _ is
            // calls action()
            // and null because we don't need about the object
            _ctx.Send(_ => action(), null);
        }

        private T OnUi<T>(Func<T> action)
        {
            SetSheet();
            if (SynchronizationContext.Current == _ctx)
                return action();


            // this line apparently initializes "result" as the "default" value of T
            // which depends : if string -> "null", if int -> 0, etc

            T result = default;


            // same as above but we actually want the result of action()
            _ctx.Send(_ => { result = action(); }, null);
            return result;
        }

        public void ToggleMainPane()
        {
            OnUi(() => _pane.Visible = !_pane.Visible);
        }

        public void SelectRange(string fp, string sp) => OnUi(() => 
            ws.Range[fp, sp].Select()
        );

        public void WriteFormula(string address, string formula) => OnUi(() => 
            ws.Range[address].Formula = formula
        );

        public Excel.Range GetCurrentSelection() => OnUi(() => 
            _app.Selection as Excel.Range
        );

        public Excel.Range GetCell(string addr) => OnUi(() =>
            ws.Range[addr]
        );

        public Excel.Range GetRange(string fp, string sp) => OnUi(() => ws.Range[fp, sp]);

        public double AddCells(Excel.Range rs) => OnUi(() => _app.WorksheetFunction.Sum(rs));

        public bool IsCellInRange(string cell, string fp, string sp)
        {
            return OnUi(() =>
            {
                var range = ws.Range[fp, sp];
                var target = ws.Range[cell];
                return _app.Application.Intersect(range, target) != null;
            });
        }
    }
}
