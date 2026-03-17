using System.Collections.Generic;
using KeePass.Plugins;
using KeePass.UI;
using KeePassLib;

namespace KeePassNetworkChecker
{
    public class NetworkStatusColumnProvider : ColumnProvider
    {
        private readonly IPluginHost m_host;
        private readonly Dictionary<string, string> m_cache = new Dictionary<string, string>();
        private readonly object m_lock = new object();

        public override string[] ColumnNames
        {
            get { return new string[] { "Net Status" }; }
        }

        public NetworkStatusColumnProvider(IPluginHost host)
        {
            m_host = host;
        }

        public override string GetCellData(string strColumnName, PwEntry pe)
        {
            if (strColumnName != "Net Status") return string.Empty;
            if (string.IsNullOrEmpty(pe.Strings.ReadSafe("URL").Trim())) return string.Empty;

            string val;
            lock (m_lock)
            {
                if (m_cache.TryGetValue(pe.Uuid.ToHexString(), out val)) return val;
            }
            return "-";
        }

        // Called only after a manual Network Check completes
        public void SetStatus(string uuid, bool isUp)
        {
            lock (m_lock) { m_cache[uuid] = isUp ? "UP" : "DOWN"; }
        }

        public void RefreshUI()
        {
            m_host.MainWindow.UpdateUI(false, null, false, null, true, null, false);
        }
    }
}
