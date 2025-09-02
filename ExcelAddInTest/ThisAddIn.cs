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
using ExcelAddInTest.Logging;

namespace ExcelAddInTest
{
    public partial class ThisAddIn
    {
        private SynchronizationContext _excelCtx;
        private ExcelApi.IExcelActions _excel;

        private Microsoft.Office.Tools.CustomTaskPane _debugPane;

        private DebugPane _debugControl;     // the UserControl instance already on your task pane
        private ILogger _log;

        private Microsoft.Office.Tools.CustomTaskPane _pane;
        private UserInterface.UserControlPane control;
        private Office.CommandBarButton _btnToggle;
        private Office.CommandBarButton _ctxToggle;

        /// <summary>
        EntityDistributor _entityDistributor;
        /// </summary>

        public Microsoft.Office.Tools.CustomTaskPane CluPane => _debugPane;

        public VoiceInterpreter Voice { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            InitializeUserInterface(); // initializam meniul 
            _excelCtx = SynchronizationContext.Current;
            
            EnsureDebugPane(); //initializam panoul de debug
            _log = new DebugPaneLogger(_debugControl);

            InitializeServices(); // initalizam speech service si clu service

            _excel = new ExcelApi.ExcelFacade(Application, _pane, _excelCtx);
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

                // addinsBar = Application.CommandBars["Add-Ins"];
            }
            catch (Exception ex)
            {
                throw new Exception("Problem on Initialize" + ex.Message);
            }
        }

        private void CellCtxToggle_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                _pane.Visible = !_pane.Visible;
                CancelDefault = true;
            }
            catch (Exception ex)
            {
                _log.Error("[CTX MENU TOGGLE ERROR]", ex);
            }
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

                _entityDistributor = new EntityDistributor(clu);

                Voice = new VoiceInterpreter(clu, _excel, new PrefixedLogger(_log, "[Speech]"), _entityDistributor);
                control.SetVoiceInterpreter(Voice);
            }   
            return Voice;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
