using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.ExcelApi
{
    public sealed class ExcelFacade : IExcelActions
    {
        private readonly Excel.Application _app;
        private readonly Microsoft.Office.Tools.CustomTaskPane _pane;
        private readonly SynchronizationContext _ctx;

        public ExcelFacade(Excel.Application app, Microsoft.Office.Tools.CustomTaskPane pane,
            SynchronizationContext ctx)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _pane = pane ?? throw new ArgumentNullException(nameof(pane));
            _ctx = ctx ?? SynchronizationContext.Current;
        }

        private void OnUi(Action action)
        {
            if (SynchronizationContext.Current == _ctx) {
                action();
                return;
            }
            _ctx.Send(_ => action(), null);
        }

        public void ToggleMainPane()
        {
            OnUi(() => _pane.Visible = !_pane.Visible);
        }

        public void SelectRange(string address) => OnUi(() => _app.Range[address].Select());
        public void WriteFormula(string address, string formula) => OnUi(() => _app.Range[address].Formula = formula);

    }
}
