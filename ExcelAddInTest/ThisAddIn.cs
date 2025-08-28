using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddInTest
{
    public partial class ThisAddIn
    {
        private SynchronizationContext _excelCtx;
        private ExcelApi.IExcelActions _excel;

        private Microsoft.Office.Tools.CustomTaskPane _debugPane;
        private DebugPane _debugControl;

        private Microsoft.Office.Tools.CustomTaskPane _pane;
        private UserInterface.UserControlPane control;
        private Office.CommandBarButton _btnToggle;
        private Office.CommandBarButton _ctxToggle;

        public Microsoft.Office.Tools.CustomTaskPane CluPane => _debugPane;

        public VoiceInterpreter Voice { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _excelCtx = SynchronizationContext.Current;
            InitializeUserInterface(); // initializam meniul 
            EnsureDebugPane(); //initializam panoul de debug
            InitializeServices(); // initalizam speech service si clu service
        }

        private void InitializeUserInterface()
        {
            control = new UserInterface.UserControlPane();

            _pane = this.CustomTaskPanes.Add(control, "Excel Voice");
            _pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _pane.Width = 580;
            _pane.Visible = true;
            CommandBar addinsBar = null;
            try
            {
                var cellMenu = Application.CommandBars["Cell"];
                if (cellMenu != null)
                {
                    _ctxToggle = (Office.CommandBarButton)cellMenu.Controls.Add(
                        Office.MsoControlType.msoControlButton, Temporary: true);
                    _ctxToggle.Caption = "Toggle Excel Voice Panel";
                    _ctxToggle.Click += CellCtxToggle_Click;
                    _ctxToggle.Style = Office.MsoButtonStyle.msoButtonCaption;
                }

                addinsBar = Application.CommandBars["Add-Ins"];
            }
            catch (Exception e) { AppendToPane("[CTX MENIU ERROR] " + e.Message); }
            _excel = new ExcelApi.ExcelFacade(Application, _pane, _excelCtx);


        }

        private void CellCtxToggle_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                _pane.Visible = !_pane.Visible;
                CancelDefault = true;
            }
            catch (Exception ex) { AppendToPane("[CTX MENIU TOGGLE ERROR] " + ex.Message); }
        }

        public void EnsureDebugPane()
        {
            _debugControl = new DebugPane();
            _debugPane = this.CustomTaskPanes.Add(_debugControl, "Debug Pane");
            _debugPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
            _debugPane.Width = 550;
            _debugPane.Height = 560;
            _debugPane.Visible = true;
            control.SetDebugPane(_debugPane);
        }



        public VoiceInterpreter InitializeServices()
        {
            if (Voice == null)
            {
                var clu = new CluService(Config.CluEndpoint, Config.CluKey, Config.CluProjectName,
                    Config.CluDeployment);
                Voice = new VoiceInterpreter(clu, _excel);
                control.SetVoiceInterpreter(Voice);
            }
            return Voice;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        public void AppendToPane(string text)
        {
            if (_debugControl == null) return;
            _debugControl.AppendText(text);
        }

        public void AppendCluLog(ExcelAddInTest.NluResult nlu)
        {
            if (nlu == null) return;

            var sb = new StringBuilder();
            sb.AppendLine("TopIntent: " + (nlu.TopIntent ?? "None"));

            if (nlu.Entities != null && nlu.Entities.Count > 0)
            {
                sb.AppendLine("Entities:");
                foreach (var ent in nlu.Entities)
                    sb.AppendLine(" - " + ent.Category + ": \"" + ent.Text + "\"");
            }
            else
            {
                sb.AppendLine("Entities: (none)");
            }

            _debugControl.AppendText(sb.ToString());
        }

        internal void AppendToInputBox(string v)
        {
            control.AppendInput(v);
        }

        internal void AppendToOutputBox(string v)
        {
            control.AppendOutput(v);
        }
    }
}
