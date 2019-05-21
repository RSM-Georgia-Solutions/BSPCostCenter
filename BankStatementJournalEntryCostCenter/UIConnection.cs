using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace BankStatementJournalEntryCostCenter
{
    class UIConnection
    {
        private static readonly Lazy<Recordset> lazy =
            new Lazy<SAPbobsCOM.Recordset>(() => (Recordset)
                xCompany.GetBusinessObject(BoObjectTypes.BoRecordset));

        private static readonly Lazy<Items> lazyItem =
            new Lazy<Items>(() => (Items)xCompany.GetBusinessObject(BoObjectTypes.oItems));

        private static readonly Lazy<SBObob> lazyBridge =new Lazy<SBObob>(() => (SBObob)xCompany.GetBusinessObject(BoObjectTypes.BoBridge));

        private static readonly Lazy<Company> lazyCompany =new Lazy<Company>(() => (Company)SAPbouiCOM.Framework
                .Application.SBO_Application.Company.GetDICompany());

        private static readonly Lazy<BarCode> lazyBarCode =
            new Lazy<BarCode>(() => (BarCode)((BarCodesService)xCompany.GetCompanyService().GetBusinessService(ServiceTypes.BarCodesService)).GetDataInterface(BarCodesServiceDataInterfaces.bsBarCode));

        public static Recordset Instance { get { return lazy.Value; } }
        public static Company xCompany { get { return lazyCompany.Value; } }
        public static SBObob xBridge { get { return lazyBridge.Value; } }
        public static Items xItem { get { return lazyItem.Value; } }
        public static BarCode xBarCode { get { return lazyBarCode.Value; } }

    }
}
