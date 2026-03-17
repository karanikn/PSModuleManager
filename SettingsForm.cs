using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using KeePass.Plugins;

namespace KeePassNetworkChecker
{
    public class SettingsForm : Form
    {
        private readonly IPluginHost m_host;

        private CheckBox      m_chkShowWindow;
        private CheckBox      m_chkResolve;
        private NumericUpDown m_numTimeout;
        private TextBox       m_txtExtraPorts;

        // Common port checkboxes: label → port number
        private static readonly int[][] CommonPorts = new int[][]
        {
            new int[] { 21   },   // FTP
            new int[] { 22   },   // SSH
            new int[] { 23   },   // Telnet
            new int[] { 25   },   // SMTP
            new int[] { 80   },   // HTTP
            new int[] { 443  },   // HTTPS
            new int[] { 554  },   // RTSP
            new int[] { 3389 },   // RDP
        };

        private static readonly string[] CommonPortLabels = new string[]
        {
            "FTP (21)", "SSH (22)", "Telnet (23)", "SMTP (25)",
            "HTTP (80)", "HTTPS (443)", "RTSP (554)", "RDP (3389)"
        };

        private CheckBox[] m_portChecks;

        public SettingsForm(IPluginHost host)
        {
            m_host = host;
            BuildUI();
        }

        private void BuildUI()
        {
            Text            = "Network Checker Options";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            StartPosition   = FormStartPosition.CenterParent;
            Font            = new Font("Segoe UI", 9f);

            int y = 12;

            // ── Title ──────────────────────────────────────────────────────
            Label lblTitle = new Label();
            lblTitle.Text     = "Network Checker Options";
            lblTitle.Font     = new Font("Segoe UI", 11f, FontStyle.Bold);
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(12, y);
            y += 34;

            // ── General ────────────────────────────────────────────────────
            m_chkShowWindow          = new CheckBox();
            m_chkShowWindow.Text     = "Show popup window when using Network Check";
            m_chkShowWindow.AutoSize = true;
            m_chkShowWindow.Location = new Point(12, y);
            m_chkShowWindow.Checked  = m_host.CustomConfig.GetBool(KeePassNetworkCheckerExt.CfgShowWindow, true);
            y += 26;

            m_chkResolve          = new CheckBox();
            m_chkResolve.Text     = "Resolve hostname to IP (DNS)";
            m_chkResolve.AutoSize = true;
            m_chkResolve.Location = new Point(12, y);
            m_chkResolve.Checked  = m_host.CustomConfig.GetBool(KeePassNetworkCheckerExt.CfgResolve, false);
            y += 30;

            // ── Timeout ────────────────────────────────────────────────────
            Label lblTimeout = new Label();
            lblTimeout.Text     = "Port scan timeout per port (ms):";
            lblTimeout.AutoSize = true;
            lblTimeout.Location = new Point(12, y + 3);

            m_numTimeout          = new NumericUpDown();
            m_numTimeout.Location = new Point(225, y);
            m_numTimeout.Size     = new Size(75, 22);
            m_numTimeout.Minimum  = 100;
            m_numTimeout.Maximum  = 10000;
            m_numTimeout.Increment = 100;
            m_numTimeout.Value    = m_host.CustomConfig.GetULong(KeePassNetworkCheckerExt.CfgTimeout, 500);

            Label lblMs = new Label();
            lblMs.Text      = "ms";
            lblMs.AutoSize  = true;
            lblMs.Location  = new Point(306, y + 3);
            y += 34;

            // ── Common ports section ───────────────────────────────────────
            GroupBox grpPorts = new GroupBox();
            grpPorts.Text     = "TCP Ports to scan";
            grpPorts.Location = new Point(12, y);
            grpPorts.Size     = new Size(490, 140);

            // Read saved enabled ports
            string savedEnabled = m_host.CustomConfig.GetString(
                KeePassNetworkCheckerExt.CfgEnabledPorts, "21,22,23,25,80,443,554,3389");
            HashSet<int> enabled = new HashSet<int>();
            foreach (string s in savedEnabled.Split(','))
            {
                int p;
                if (int.TryParse(s.Trim(), out p)) enabled.Add(p);
            }

            m_portChecks = new CheckBox[CommonPorts.Length];
            int cx = 10; int cy = 20;
            for (int i = 0; i < CommonPorts.Length; i++)
            {
                CheckBox chk = new CheckBox();
                chk.Text     = CommonPortLabels[i];
                chk.Tag      = CommonPorts[i][0];
                chk.AutoSize = true;
                chk.Location = new Point(cx, cy);
                chk.Checked  = enabled.Contains(CommonPorts[i][0]);
                m_portChecks[i] = chk;
                grpPorts.Controls.Add(chk);

                if ((i + 1) % 4 == 0) { cx = 10; cy += 26; }
                else cx += 118;
            }

            // Extra ports text field
            Label lblExtra = new Label();
            lblExtra.Text     = "Additional ports (comma-separated):";
            lblExtra.AutoSize = true;
            lblExtra.Location = new Point(10, cy + 28);
            grpPorts.Controls.Add(lblExtra);

            m_txtExtraPorts          = new TextBox();
            m_txtExtraPorts.Location = new Point(10, cy + 46);
            m_txtExtraPorts.Size     = new Size(465, 22);
            m_txtExtraPorts.Text     = m_host.CustomConfig.GetString(KeePassNetworkCheckerExt.CfgExtraPorts, "");
            grpPorts.Controls.Add(m_txtExtraPorts);

            grpPorts.Size = new Size(490, cy + 76);
            y += grpPorts.Height + 12;

            // ── Hint ───────────────────────────────────────────────────────
            Label lblHint = new Label();
            lblHint.Text      = "To show the status column: View \u2192 Configure Columns \u2192 enable 'Net Status'";
            lblHint.ForeColor = SystemColors.GrayText;
            lblHint.Location  = new Point(12, y);
            lblHint.Size      = new Size(490, 18);
            y += 28;

            // ── Buttons ────────────────────────────────────────────────────
            Button btnOk = new Button();
            btnOk.Text         = "OK";
            btnOk.Size         = new Size(75, 26);
            btnOk.Location     = new Point(416, y);
            btnOk.DialogResult = DialogResult.OK;
            btnOk.Click       += OnOkClick;

            Button btnCancel = new Button();
            btnCancel.Text         = "Cancel";
            btnCancel.Size         = new Size(75, 26);
            btnCancel.Location     = new Point(332, y);
            btnCancel.DialogResult = DialogResult.Cancel;

            ClientSize = new Size(514, y + 44);

            Controls.Add(lblTitle);
            Controls.Add(m_chkShowWindow);
            Controls.Add(m_chkResolve);
            Controls.Add(lblTimeout);
            Controls.Add(m_numTimeout);
            Controls.Add(lblMs);
            Controls.Add(grpPorts);
            Controls.Add(lblHint);
            Controls.Add(btnOk);
            Controls.Add(btnCancel);
        }

        private void OnOkClick(object sender, EventArgs e)
        {
            m_host.CustomConfig.SetBool(KeePassNetworkCheckerExt.CfgShowWindow, m_chkShowWindow.Checked);
            m_host.CustomConfig.SetBool(KeePassNetworkCheckerExt.CfgResolve,    m_chkResolve.Checked);
            m_host.CustomConfig.SetULong(KeePassNetworkCheckerExt.CfgTimeout,   (ulong)m_numTimeout.Value);

            // Save checked common ports
            List<string> enabledList = new List<string>();
            foreach (CheckBox chk in m_portChecks)
                if (chk.Checked) enabledList.Add(((int)chk.Tag).ToString());
            m_host.CustomConfig.SetString(KeePassNetworkCheckerExt.CfgEnabledPorts,
                string.Join(",", enabledList.ToArray()));

            // Save extra ports
            m_host.CustomConfig.SetString(KeePassNetworkCheckerExt.CfgExtraPorts,
                m_txtExtraPorts.Text.Trim());

            Close();
        }
    }
}
