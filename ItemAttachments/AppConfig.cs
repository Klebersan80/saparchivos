using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ItemAttachments
{
    public class AppConfig
    {

        public string SAPBobCompany { get; set; }
        public string SAPBobServer { get; set; }
        public string SAPBobUsername { get; set; }
        public string SAPBobPassword { get; set; }
        public bool SAPBobTrusted { get; set; }

        public void getSettings()
        {
            // App.config Settings
            System.Configuration.AppSettingsReader appReader = new System.Configuration.AppSettingsReader();

            // For Basic Reads from SAP BOB
            SAPBobCompany = appReader.GetValue("SAPBobCompany", typeof(string)).ToString();
            SAPBobServer = appReader.GetValue("SAPBobServer", typeof(string)).ToString();
            SAPBobUsername = appReader.GetValue("SAPBobUsername", typeof(string)).ToString();
            SAPBobPassword = appReader.GetValue("SAPBobPassword", typeof(string)).ToString();
            SAPBobTrusted = (bool)appReader.GetValue("SAPBobTrusted", typeof(bool));
        }
    }
}
