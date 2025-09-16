using Microsoft.Office.Core;
using System;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using ExcelAddInTest.Logging;
using ExcelAddInTest.Nlu;
using ExcelAddInTest.ExcelApi;
using ExcelAddInTest.UserInterface;
using Microsoft.Office.Tools;
using ExcelAddInTest.Speech;

namespace ExcelAddInTest
{
    public partial class ThisAddIn
    {
        private CluService _clu;
        private CommandExecutor _executor;
        private IIntentRouter _intentRouter;

        private SynchronizationContext _excelCtx;
        private IExcelActions _excel;

        private Microsoft.Office.Tools.CustomTaskPane _debugPane;

        private DebugPane _debugControl;     // the UserControl instance already on your task pane
        private ILogger _log;

        private Microsoft.Office.Tools.CustomTaskPane _pane;
        private UserInterface.UserControlPane control;
        private CommandBarButton _btnToggle;
        private CommandBarButton _ctxToggle;

        private Microsoft.Office.Tools.CustomTaskPane _settingsPane;
        private SettingsPane _settingsControl;


        EntityDistributor _entityDistributor;
        private VoiceListenOptions VoiceListenOptions = new VoiceListenOptions()
        {
            Mode = ListenMode.Continuous,
            AutoStopAfter = TimeSpan.FromSeconds(15),
            MaxDuration = TimeSpan.FromSeconds(30)
            // Language/Initial/End silence are fixed in the class
        };
        public Microsoft.Office.Tools.CustomTaskPane CluPane => _debugPane;

        public VoiceInterpreter Voice { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            InitializeUserInterface(); // initializam meniul 
            _excelCtx = SynchronizationContext.Current;
            
            EnsureDebugPane(); //initializam panoul de debug
            EnsureSettingsPane(); //initializam panoul de setari
            
            _log = new DebugPaneLogger(_debugControl);
            _excel = new ExcelApi.ExcelFacade(Application, _pane, _excelCtx);
            
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

                // addinsBar = Application.CommandBars["Add-Ins"];
            }
            catch (Exception ex)
            {
                throw new Exception("Problem on Initialize" + ex.Message);
            }
        }

        //add a toggle button in the right click meniu of a cell
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

        public void EnsureSettingsPane()
        {
            if (_settingsPane != null) return;   // avoid double-creating
            _settingsControl = new SettingsPane();

            // 1) sync UI from the current shared options
            _settingsControl.LoadOptions(VoiceListenOptions);

            // 2) when user changes settings in the pane, update the shared instance
            //    and push live updates into the recognizer
            _settingsControl.OptionsChanged += opts =>
            {
                VoiceListenOptions = opts;           // keep the one-and-only instance up to date
                Voice?.UpdateOptions(VoiceListenOptions);
            };

            _settingsPane = this.CustomTaskPanes.Add(_settingsControl, "Settings Pane");
            _settingsPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _settingsPane.Width = 340;
            // _settingsPane.Height = 200;
            _settingsPane.Visible = false;
            control.SetSettingsPane(_settingsPane);
            control.SetOptionsRef(VoiceListenOptions);
        }

        public VoiceInterpreter InitializeServices()
        {
            _executor = new ExcelApi.CommandExecutor(_excel, new PrefixedLogger(_log, "[CmdExec]"));
            _intentRouter = new IntentRouter();
            if (Voice == null)
            {
                _clu = new CluService(Config.CluEndpoint, Config.CluKey, Config.CluProjectName,
                    Config.CluDeployment);

                _entityDistributor = new EntityDistributor(_clu);

                Voice = new VoiceInterpreter(_clu, _executor, new PrefixedLogger(_log, "[Speech]"), _entityDistributor, _intentRouter);
                control.SetVoiceInterpreter(Voice);
            }   
            return Voice;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
    }
}
