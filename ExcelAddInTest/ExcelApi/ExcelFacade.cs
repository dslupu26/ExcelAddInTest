using ExcelAddInTest.ExcelApi.Commands;
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
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

        // Simple A1 validator to fail fast before COM
        private static readonly Regex RxA1 =
            new Regex(@"^[A-Za-z]{1,3}\d{1,7}$", RegexOptions.Compiled);

        public ExcelFacade(Excel.Application app, Microsoft.Office.Tools.CustomTaskPane pane, SynchronizationContext ctx)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _pane = pane ?? throw new ArgumentNullException(nameof(pane));
            _ctx = ctx ?? SynchronizationContext.Current;
            ws = _app.ActiveSheet;
        }

        public void RefreshActiveSheet()
        {
            ws = _app.ActiveSheet as Excel.Worksheet
                  ?? throw new InvalidOperationException("No active worksheet.");
        }
        // this returns nothing
        // therefore it's only good when you don't want to get back *something*
        // kept it here because I don't know if the code explodes without it

        private static void ValidateA1(string a1, string paramName)
        {
            if (string.IsNullOrWhiteSpace(a1))
                throw new ArgumentNullException(paramName);
            if (!RxA1.IsMatch(a1))
                throw new ArgumentException($"Invalid A1 address: '{a1}'", paramName);
        }

        // Marshal actions/funcs to Excel UI thread
        // add this overload next to your existing OnUi<T>
        private void OnUi(Action action)
        {
            RefreshActiveSheet();
            if (SynchronizationContext.Current == _ctx) action();
            else _ctx.Send(_ => action(), null);
        }

        private T OnUi<T>(Func<T> action)
        {
            RefreshActiveSheet();
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

        public void SelectRange(string firstA1, string secondA1) =>
            OnUi(() =>
            {
                ValidateA1(firstA1, nameof(firstA1));
                ValidateA1(secondA1, nameof(secondA1));
                try
                {
                    var rng = ws.Range[firstA1, secondA1]; // handles reversed order
                    rng.Select();
                }
                catch (COMException ex)
                {
                    throw new CommandExecutionException(
                        $"Failed to select range {firstA1}:{secondA1}.", ex);
                }
            });

        public void WriteInCell(string cell, string text) =>
            OnUi(() =>
            {
                ValidateA1(cell, nameof(cell));
                try
                {
                    var rng = ws.Range[cell];
                    var s = (text ?? string.Empty).Trim();

                    if (s.StartsWith("=", StringComparison.Ordinal))
                    {
                        rng.Formula = s;
                        return;
                    }

                    if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
                        rng.Value2 = d;
                    else
                        rng.Value2 = s;
                }
                catch (COMException ex)
                {
                    throw new CommandExecutionException($"Excel write failed at '{cell}'.", ex);
                }
            });

        public void WriteFormula(string address, string formula) =>
            OnUi(() =>
            {
                ValidateA1(address, nameof(address));
                try
                {
                    ws.Range[address].Formula = (formula ?? string.Empty).Trim();
                }
                catch (COMException ex)
                {
                    throw new CommandExecutionException($"Failed to set formula at '{address}'.", ex);
                }
            });

        public Excel.Range GetCurrentSelection() =>
            OnUi(() =>
            {
                try { return _app.Selection as Excel.Range; }
                catch (COMException ex)
                {
                    throw new CommandExecutionException("Failed to get current selection.", ex);
                }
            });

        public Excel.Range GetCell(string addr) =>
            OnUi(() =>
            {
                ValidateA1(addr, nameof(addr));
                try { return ws.Range[addr]; }
                catch (COMException ex)
                {
                    throw new CommandExecutionException($"Failed to get cell '{addr}'.", ex);
                }
            });

        public Excel.Range GetRange(string fp, string sp) =>
            OnUi(() =>
            {
                ValidateA1(fp, nameof(fp));
                ValidateA1(sp, nameof(sp));
                try { return ws.Range[fp, sp]; }
                catch (COMException ex)
                {
                    throw new CommandExecutionException($"Failed to get range {fp}:{sp}.", ex);
                }
            });

        public double AddCells(Excel.Range rs) =>
            OnUi(() =>
            {
                if (rs == null) throw new ArgumentNullException(nameof(rs));
                try { return _app.WorksheetFunction.Sum(rs); }
                catch (COMException ex)
                {
                    throw new CommandExecutionException("SUM failed on the provided range.", ex);
                }
            });

        public bool IsCellInRange(string cell, string fp, string sp) =>
            OnUi(() =>
            {
                ValidateA1(cell, nameof(cell));
                ValidateA1(fp, nameof(fp));
                ValidateA1(sp, nameof(sp));
                try
                {
                    var range = ws.Range[fp, sp];
                    var target = ws.Range[cell];
                    return _app.Application.Intersect(range, target) != null;
                }
                catch (COMException ex)
                {
                    throw new CommandExecutionException($"Intersect check failed for {cell} in {fp}:{sp}.", ex);
                }
            });
    }
}
