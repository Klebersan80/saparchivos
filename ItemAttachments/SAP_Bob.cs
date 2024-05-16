using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ItemAttachments
{
    public class SAP_Bob
    {
        public SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

        private readonly string _SAPBobCompany;
        private readonly string _SAPBobServer;
        private readonly string _SAPBobUsername;
        private readonly string _SAPBobPassword;
        private readonly bool _SAPBobTrusted;

        public SAP_Bob()
        {
            AppConfig Config = new AppConfig();
            Config.getSettings();
            _SAPBobCompany = Config.SAPBobCompany;
            _SAPBobServer = Config.SAPBobServer;
            _SAPBobUsername = Config.SAPBobUsername;
            _SAPBobPassword = Config.SAPBobPassword;
            _SAPBobTrusted = Config.SAPBobTrusted;
        }

        public SAPbobsCOM.Company _getCompanyConnection()
        {

            oCompany = new SAPbobsCOM.Company();
            oCompany.Server = _SAPBobServer;
            oCompany.CompanyDB = _SAPBobCompany;
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
            oCompany.UserName = _SAPBobUsername;
            oCompany.Password = _SAPBobPassword;
            oCompany.UseTrusted = _SAPBobTrusted;

            oCompany.Connect();

            return oCompany;
        }
    }
}
