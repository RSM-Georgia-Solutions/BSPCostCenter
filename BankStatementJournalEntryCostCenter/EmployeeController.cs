using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace BankStatementJournalEntryCostCenter
{
    class EmployeeController
    {
        public static EmployeesInfo Get(string FederalTaxId)
        {
            var oEmployee = (EmployeesInfo)UIConnection.xCompany.GetBusinessObject(BoObjectTypes.oEmployeesInfo);
            var rs = UIConnection.xBridge.GetObjectKeyBySingleValue(BoObjectTypes.oEmployeesInfo, "IdNumber", FederalTaxId, BoQueryConditions.bqc_Equal);
            if (rs.RecordCount != 0)
            {
                int Code = int.Parse(rs.Fields.Item("empID").Value.ToString());
                oEmployee = (EmployeesInfo)UIConnection.xCompany.GetBusinessObject(BoObjectTypes.oEmployeesInfo);
                var res = oEmployee.GetByKey(Code);
                if (res == true)
                {
                    return oEmployee;
                }
            }
            throw new Exception("Employee not found");
        }
    }
}
