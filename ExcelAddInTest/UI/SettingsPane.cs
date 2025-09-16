using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace ExcelAddInTest
{
    public partial class SettingsPane : UserControl
    {
        // the shared options instance
        private VoiceListenOptions _opts = new VoiceListenOptions();

        // notify ThisAddIn/VoiceInterpreter when something changes
        public event Action<VoiceListenOptions> OptionsChanged;

        public SettingsPane()
        {
            InitializeComponent();

            // wire UI events (designer contains the controls)
            chkContinuous.CheckedChanged += chkContinuous_CheckedChanged;
            chkSingle.CheckedChanged += chkSingle_CheckedChanged;
            tbSeconds.ValueChanged += tbSeconds_ValueChanged;

            // show current value in the label
            lblSeconds.Text = $"Listen for: {tbSeconds.Value} s";

            // don’t run runtime wiring in the designer
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;

            // initial sync (uses defaults until LoadOptions is called)
            SyncToUi();
        }

        /// <summary>Push the shared options into the UI (call this before showing the pane).</summary>
        public void LoadOptions(VoiceListenOptions opts)
        {
            _opts = opts ?? new VoiceListenOptions();
            SyncToUi();
        }

        /// <summary>Reflects _opts into the checkboxes + slider.</summary>
        private void SyncToUi()
        {
            // mode
            chkContinuous.Checked = (_opts.Mode == ListenMode.Continuous);
            chkSingle.Checked = (_opts.Mode == ListenMode.SingleUtterance);

            // seconds (clamp to 1..30)
            var secs = (int)Math.Max(1, Math.Min(30,
                (_opts.AutoStopAfter ?? TimeSpan.FromSeconds(15)).TotalSeconds));
            if (secs < tbSeconds.Minimum) secs = tbSeconds.Minimum;
            if (secs > tbSeconds.Maximum) secs = tbSeconds.Maximum;

            // avoid bouncing ValueChanged during sync
            if (tbSeconds.Value != secs) tbSeconds.Value = secs;
            lblSeconds.Text = $"Listen for: {secs} s";

            tbSeconds.Enabled = chkContinuous.Checked;
        }

        private void chkContinuous_CheckedChanged(object sender, EventArgs e)
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            if (!chkContinuous.Checked) return;

            // mutually exclusive with Single
            if (chkSingle.Checked) chkSingle.Checked = false;

            _opts.Mode = ListenMode.Continuous;
            _opts.AutoStopAfter = TimeSpan.FromSeconds(tbSeconds.Value);
            _opts.MaxDuration = TimeSpan.FromSeconds(30); // absolute cap
            tbSeconds.Enabled = true;

            OptionsChanged?.Invoke(_opts);
        }

        private void chkSingle_CheckedChanged(object sender, EventArgs e)
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            if (!chkSingle.Checked) return;

            if (chkContinuous.Checked) chkContinuous.Checked = false;

            _opts.Mode = ListenMode.SingleUtterance;
            // engine stops after ~3s silence via EndSilenceTimeoutMs=3000; keep a 30s safety cap
            _opts.AutoStopAfter = TimeSpan.FromSeconds(30);
            _opts.MaxDuration = TimeSpan.FromSeconds(30);
            tbSeconds.Enabled = false;

            OptionsChanged?.Invoke(_opts);
        }

        private void tbSeconds_ValueChanged(object sender, EventArgs e)
        {
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;

            lblSeconds.Text = $"Listen for: {tbSeconds.Value} s";

            if (chkContinuous.Checked && _opts != null)
            {
                _opts.AutoStopAfter = TimeSpan.FromSeconds(tbSeconds.Value);
                OptionsChanged?.Invoke(_opts);
            }
        }
    }
}

