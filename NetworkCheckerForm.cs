using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Windows.Forms;
using KeePassLib;

namespace KeePassNetworkChecker
{
    public class NetworkCheckerForm : Form
    {
        private readonly PwEntry[]                _entries;
        private readonly KeePassNetworkCheckerExt _plugin;
        private DataGridView     _grid;
        private Button           _btnRefresh;
        private Button           _btnClose;
        private Button           _btnExport;
        private Label            _lblStatus;
        private ComboBox         _cmbFilter;
        private List<CheckResult> _lastResults = new List<CheckResult>();

        private class CheckResult
        {
            public PwEntry Entry;
            public string  Device;
            public string  URL;
            public string  ResolvedIP;
            public string  Ping;
            public string  OpenPorts;
            public string  Web;
            public bool    PingOk;
            public bool    PortOk;
            public bool    WebOk;
            public bool    AnyOk { get { return PingOk || PortOk || WebOk; } }
        }

        public NetworkCheckerForm(PwEntry[] entries, KeePassNetworkCheckerExt plugin)
        {
            _entries = entries;
            _plugin  = plugin;
            BuildUI();
            SetupWorker();
        }

        private BackgroundWorker _worker;

        private void BuildUI()
        {
            Text          = "Network Checker";
            Size          = new Size(1050, 500);
            MinimumSize   = new Size(850, 400);
            StartPosition = FormStartPosition.CenterParent;
            Font          = new Font("Segoe UI", 9f);

            // ── Title ──────────────────────────────────────────────────────
            Label lblTitle = new Label();
            lblTitle.Text     = "Network Checker";
            lblTitle.Font     = new Font("Segoe UI", 13f, FontStyle.Bold);
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(12, 12);

            _lblStatus = new Label();
            _lblStatus.Text     = "Ready.";
            _lblStatus.AutoSize = true;
            _lblStatus.Location = new Point(12, 46);

            // ── Filter ─────────────────────────────────────────────────────
            Label lblFilter = new Label();
            lblFilter.Text     = "Show:";
            lblFilter.AutoSize = true;
            lblFilter.Location = new Point(12, 70);

            _cmbFilter = new ComboBox();
            _cmbFilter.Location       = new Point(52, 67);
            _cmbFilter.Size           = new Size(100, 22);
            _cmbFilter.DropDownStyle  = ComboBoxStyle.DropDownList;
            _cmbFilter.Items.AddRange(new object[] { "All", "UP only", "DOWN only" });
            _cmbFilter.SelectedIndex  = 0;
            _cmbFilter.SelectedIndexChanged += (s, e) => ApplyFilter();

            // ── Grid ───────────────────────────────────────────────────────
            _grid = new DataGridView();
            _grid.Location              = new Point(12, 96);
            _grid.Anchor                = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _grid.Size                  = new Size(1010, 320);
            _grid.BorderStyle           = BorderStyle.FixedSingle;
            _grid.CellBorderStyle       = DataGridViewCellBorderStyle.SingleHorizontal;
            _grid.SelectionMode         = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect           = false;
            _grid.ReadOnly              = true;
            _grid.AllowUserToAddRows    = false;
            _grid.AllowUserToResizeRows = false;
            _grid.RowHeadersVisible     = false;
            _grid.AutoSizeColumnsMode   = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.RowTemplate.Height    = 24;
            _grid.SortCompare          += Grid_SortCompare;

            var colDevice = new DataGridViewTextBoxColumn { Name = "Device",     HeaderText = "Device",     FillWeight = 16, SortMode = DataGridViewColumnSortMode.Automatic };
            var colURL    = new DataGridViewTextBoxColumn { Name = "URL",        HeaderText = "URL",        FillWeight = 18, SortMode = DataGridViewColumnSortMode.Automatic };
            var colIP     = new DataGridViewTextBoxColumn { Name = "ResolvedIP", HeaderText = "IP",         FillWeight = 10, SortMode = DataGridViewColumnSortMode.Automatic };
            var colPing   = new DataGridViewTextBoxColumn { Name = "Ping",       HeaderText = "Ping",       FillWeight = 9,  SortMode = DataGridViewColumnSortMode.Automatic };
            var colPorts  = new DataGridViewTextBoxColumn { Name = "OpenPorts",  HeaderText = "Open Ports", FillWeight = 30, SortMode = DataGridViewColumnSortMode.Automatic };
            var colWeb    = new DataGridViewTextBoxColumn { Name = "Web",        HeaderText = "HTTP",       FillWeight = 8,  SortMode = DataGridViewColumnSortMode.Automatic };
            var colStatus = new DataGridViewTextBoxColumn { Name = "Overall",    HeaderText = "Status",     FillWeight = 9,  SortMode = DataGridViewColumnSortMode.Automatic };

            _grid.Columns.AddRange(new DataGridViewColumn[] { colDevice, colURL, colIP, colPing, colPorts, colWeb, colStatus });
            _grid.CellFormatting      += Grid_CellFormatting;
            _grid.MouseClick          += Grid_MouseClick;

            // ── Bottom panel ───────────────────────────────────────────────
            _btnRefresh = new Button();
            _btnRefresh.Text    = "Refresh";
            _btnRefresh.Size    = new Size(80, 26);
            _btnRefresh.Enabled = false;
            _btnRefresh.Click  += OnRefreshClick;

            _btnExport = new Button();
            _btnExport.Text    = "Export CSV";
            _btnExport.Size    = new Size(90, 26);
            _btnExport.Enabled = false;
            _btnExport.Click  += OnExportClick;

            _btnClose = new Button();
            _btnClose.Text         = "Close";
            _btnClose.Size         = new Size(80, 26);
            _btnClose.DialogResult = DialogResult.Cancel;

            FlowLayoutPanel pnl = new FlowLayoutPanel();
            pnl.FlowDirection = FlowDirection.RightToLeft;
            pnl.Dock          = DockStyle.Bottom;
            pnl.Height        = 40;
            pnl.Padding       = new Padding(4);
            pnl.Controls.Add(_btnClose);
            pnl.Controls.Add(_btnRefresh);
            pnl.Controls.Add(_btnExport);

            Controls.Add(lblTitle);
            Controls.Add(_lblStatus);
            Controls.Add(lblFilter);
            Controls.Add(_cmbFilter);
            Controls.Add(_grid);
            Controls.Add(pnl);

            Resize += (s, e) => _grid.Size = new Size(ClientSize.Width - 24, ClientSize.Height - 156);
            Shown  += (s, e) => StartChecks();
        }

        // ── Sorting ────────────────────────────────────────────────────────
        private void Grid_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            e.SortResult  = string.Compare(
                e.CellValue1 != null ? e.CellValue1.ToString() : "",
                e.CellValue2 != null ? e.CellValue2.ToString() : "");
            e.Handled     = true;
        }

        // ── Cell coloring ──────────────────────────────────────────────────
        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.Value == null) return;
            string col = _grid.Columns[e.ColumnIndex].Name;
            string val = e.Value.ToString();

            if (col == "Overall")
            {
                e.CellStyle.ForeColor = val == "UP" ? Color.Green : Color.Red;
                e.CellStyle.Font      = new Font(_grid.Font, FontStyle.Bold);
            }
            else if (col == "Ping")
                e.CellStyle.ForeColor = val.StartsWith("TIMEOUT") || val == "ERR" ? Color.Red : Color.Green;
            else if (col == "OpenPorts")
                e.CellStyle.ForeColor = val == "none" || val == "N/A" ? Color.Red : Color.Green;
            else if (col == "Web")
                e.CellStyle.ForeColor = val.StartsWith("ERR") ? Color.Red : Color.Green;
        }

        // ── Right-click context menu ───────────────────────────────────────
        private void Grid_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            var hit = _grid.HitTest(e.X, e.Y);
            if (hit.RowIndex < 0) return;
            _grid.ClearSelection();
            _grid.Rows[hit.RowIndex].Selected = true;

            DataGridViewRow row = _grid.Rows[hit.RowIndex];
            string ip  = row.Cells["ResolvedIP"].Value != null ? row.Cells["ResolvedIP"].Value.ToString() : "";
            string url = row.Cells["URL"].Value != null ? row.Cells["URL"].Value.ToString() : "";

            ContextMenuStrip ctx = new ContextMenuStrip();

            ToolStripMenuItem miCopyIP = new ToolStripMenuItem("Copy IP");
            miCopyIP.Enabled = ip != "-" && ip != "" && ip != "ERR";
            miCopyIP.Click  += delegate(object s, EventArgs ev)
            {
                try { Clipboard.SetText(ip); _lblStatus.Text = "Copied: " + ip; } catch { }
            };

            ToolStripMenuItem miCopyURL = new ToolStripMenuItem("Copy URL");
            miCopyURL.Enabled = !string.IsNullOrEmpty(url);
            miCopyURL.Click  += delegate(object s, EventArgs ev)
            {
                try { Clipboard.SetText(url); _lblStatus.Text = "Copied: " + url; } catch { }
            };

            ToolStripMenuItem miOpenURL = new ToolStripMenuItem("Open URL in browser");
            miOpenURL.Enabled = !string.IsNullOrEmpty(url);
            miOpenURL.Click  += delegate(object s, EventArgs ev)
            {
                try
                {
                    string fullUrl = url.Contains("://") ? url : "http://" + url;
                    Process.Start(fullUrl);
                }
                catch { }
            };

            ToolStripMenuItem miPingAgain = new ToolStripMenuItem("Ping again");
            miPingAgain.Click += delegate(object s, EventArgs ev)
            {
                string host = url;
                try { Uri uri = new Uri(url.Contains("://") ? url : "http://" + url); host = uri.Host; } catch { }
                var res = DoPing(host);
                row.Cells["Ping"].Value = res.Item1 ? res.Item2 + " ms" : "TIMEOUT";
                _lblStatus.Text = "Pinged " + host + ": " + row.Cells["Ping"].Value;
            };

            ctx.Items.Add(miCopyIP);
            ctx.Items.Add(miCopyURL);
            ctx.Items.Add(miOpenURL);
            ctx.Items.Add(new ToolStripSeparator());
            ctx.Items.Add(miPingAgain);
            ctx.Show(_grid, e.Location);
        }

        // ── Filter ─────────────────────────────────────────────────────────
        private void ApplyFilter()
        {
            if (_lastResults.Count == 0) return;
            string filter = _cmbFilter.SelectedItem.ToString();
            _grid.Rows.Clear();
            foreach (CheckResult r in _lastResults)
            {
                if (filter == "UP only"   && !r.AnyOk) continue;
                if (filter == "DOWN only" &&  r.AnyOk) continue;
                AddRow(r);
            }
        }

        private void AddRow(CheckResult r)
        {
            int idx = _grid.Rows.Add();
            DataGridViewRow row = _grid.Rows[idx];
            row.Cells["Device"].Value     = r.Device;
            row.Cells["URL"].Value        = r.URL;
            row.Cells["ResolvedIP"].Value = r.ResolvedIP ?? "-";
            row.Cells["Ping"].Value       = r.Ping       ?? "N/A";
            row.Cells["OpenPorts"].Value  = r.OpenPorts  ?? "N/A";
            row.Cells["Web"].Value        = r.Web        ?? "N/A";
            row.Cells["Overall"].Value    = r.AnyOk ? "UP" : "DOWN";
        }

        // ── Export CSV ─────────────────────────────────────────────────────
        private void OnExportClick(object sender, EventArgs e)
        {
            if (_lastResults.Count == 0) return;
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter   = "CSV files (*.csv)|*.csv";
            dlg.FileName  = "NetworkCheck_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".csv";
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Device,URL,IP,Ping,Open Ports,HTTP,Status");
                foreach (CheckResult r in _lastResults)
                {
                    sb.AppendLine(string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\"",
                        r.Device, r.URL, r.ResolvedIP ?? "-",
                        r.Ping ?? "N/A", r.OpenPorts ?? "N/A",
                        r.Web ?? "N/A", r.AnyOk ? "UP" : "DOWN"));
                }
                File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                _lblStatus.Text = "Exported: " + dlg.FileName;
            }
            catch (Exception ex) { MessageBox.Show("Export failed: " + ex.Message); }
        }

        private void OnRefreshClick(object sender, EventArgs e) { StartChecks(); }

        // ── BackgroundWorker ───────────────────────────────────────────────
        private void SetupWorker()
        {
            _worker = new BackgroundWorker();
            _worker.WorkerReportsProgress      = true;
            _worker.WorkerSupportsCancellation = true;
            _worker.DoWork             += Worker_DoWork;
            _worker.ProgressChanged    += Worker_ProgressChanged;
            _worker.RunWorkerCompleted += Worker_Completed;
        }

        private void StartChecks()
        {
            if (_worker.IsBusy) return;
            _btnRefresh.Enabled = false;
            _btnExport.Enabled  = false;
            _cmbFilter.SelectedIndex = 0;
            _lblStatus.Text     = "Checking...";
            _grid.Rows.Clear();
            _lastResults.Clear();
            _worker.RunWorkerAsync(_entries);
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            PwEntry[] entries   = (PwEntry[])e.Argument;
            int[]     ports     = _plugin.GetPorts();
            int       timeout   = _plugin.GetTimeout();
            bool      doResolve = _plugin.GetResolve();
            Stopwatch sw        = Stopwatch.StartNew();

            List<CheckResult> results = new List<CheckResult>();

            foreach (PwEntry entry in entries)
            {
                string title  = entry.Strings.ReadSafe("Title");
                string url    = entry.Strings.ReadSafe("URL").Trim();
                CheckResult r = new CheckResult { Entry = entry, Device = title, URL = url, ResolvedIP = "-" };

                if (string.IsNullOrEmpty(url))
                {
                    r.Ping = "N/A"; r.OpenPorts = "N/A"; r.Web = "N/A";
                    results.Add(r);
                    _worker.ReportProgress(results.Count, title);
                    continue;
                }

                string fullUrl = url.Contains("://") ? url : "http://" + url;
                string host    = url;
                try { Uri uri = new Uri(fullUrl); host = uri.Host; } catch { }

                // DNS
                if (doResolve)
                {
                    try
                    {
                        IPHostEntry he = Dns.GetHostEntry(host);
                        if (he.AddressList.Length > 0) r.ResolvedIP = he.AddressList[0].ToString();
                    }
                    catch { r.ResolvedIP = "ERR"; }
                }

                // Ping
                var pingRes = DoPing(host, timeout);
                r.PingOk = pingRes.Item1;
                r.Ping   = pingRes.Item1 ? pingRes.Item2 + " ms" : "TIMEOUT";

                // Ports
                List<int> open = new List<int>();
                foreach (int port in ports)
                {
                    try
                    {
                        using (TcpClient client = new TcpClient())
                        {
                            IAsyncResult ar = client.BeginConnect(host, port, null, null);
                            if (ar.AsyncWaitHandle.WaitOne(timeout) && client.Connected)
                            {
                                try { client.EndConnect(ar); } catch { }
                                open.Add(port);
                            }
                        }
                    }
                    catch { }
                }
                r.PortOk    = open.Count > 0;
                r.OpenPorts = open.Count > 0 ? string.Join(", ", open.ToArray()) : "none";

                // HTTP
                try
                {
                    ServicePointManager.ServerCertificateValidationCallback =
                        delegate(object s,
                            System.Security.Cryptography.X509Certificates.X509Certificate c,
                            System.Security.Cryptography.X509Certificates.X509Chain ch,
                            System.Net.Security.SslPolicyErrors er) { return true; };
                    ServicePointManager.SecurityProtocol =
                        (SecurityProtocolType)3072 | (SecurityProtocolType)768;

                    HttpWebRequest req    = (HttpWebRequest)WebRequest.Create(fullUrl);
                    req.Timeout           = timeout * 3;
                    req.AllowAutoRedirect = true;
                    req.Method            = "GET";
                    req.UserAgent         = "KeePassNetworkChecker/1.2";

                    using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                    {
                        int code = (int)resp.StatusCode;
                        r.WebOk = code >= 200 && code < 400;
                        r.Web   = code.ToString();
                    }
                }
                catch (WebException ex)
                {
                    HttpWebResponse err = ex.Response as HttpWebResponse;
                    r.Web = err != null ? "ERR " + (int)err.StatusCode : "ERR";
                }
                catch { r.Web = "ERR"; }

                results.Add(r);
                _worker.ReportProgress(results.Count, title);
            }

            sw.Stop();
            e.Result = new object[] { results, sw.Elapsed };
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            _lblStatus.Text = "Checking " + (string)e.UserState + "...";
        }

        private void Worker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                _lblStatus.Text     = "Error: " + e.Error.Message;
                _btnRefresh.Enabled = true;
                return;
            }

            object[]          data    = (object[])e.Result;
            List<CheckResult> results = (List<CheckResult>)data[0];
            TimeSpan          elapsed = (TimeSpan)data[1];

            _lastResults = results;
            _grid.Rows.Clear();

            foreach (CheckResult r in results)
            {
                AddRow(r);
                if (_plugin != null && _plugin.ColProvider != null && r.Entry != null)
                    _plugin.ColProvider.SetStatus(r.Entry.Uuid.ToHexString(), r.AnyOk);
            }

            if (_plugin != null && _plugin.ColProvider != null)
                _plugin.ColProvider.RefreshUI();

            string elapsedStr = elapsed.TotalSeconds >= 60
                ? string.Format("{0}m {1}s", (int)elapsed.TotalMinutes, elapsed.Seconds)
                : string.Format("{0:F1}s", elapsed.TotalSeconds);

            _lblStatus.Text     = string.Format("Last checked: {0}  -  {1} device(s)  -  Scan time: {2}",
                DateTime.Now.ToString("HH:mm:ss"), results.Count, elapsedStr);
            _btnRefresh.Enabled = true;
            _btnExport.Enabled  = results.Count > 0;
        }

        // ── Network helpers ────────────────────────────────────────────────
        private static Tuple<bool, long> DoPing(string host, int timeout = 3000)
        {
            try
            {
                using (Ping ping = new Ping())
                {
                    PingReply reply = ping.Send(host, timeout);
                    if (reply != null && reply.Status == IPStatus.Success)
                        return Tuple.Create(true, reply.RoundtripTime);
                }
            }
            catch { }
            return Tuple.Create(false, 0L);
        }
    }
}
