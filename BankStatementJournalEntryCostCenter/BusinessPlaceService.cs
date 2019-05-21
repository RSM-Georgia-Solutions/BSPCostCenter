using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankStatementJournalEntryCostCenter
{
    public static class BusinessPlaceService
    {
        public static BusinessPlaces GetBusinessPlace(int businessPlaceId)
        {
            var oBusinessPlace = (BusinessPlaces)UIConnection.xCompany.GetBusinessObject(BoObjectTypes.oBusinessPlaces);
            oBusinessPlace.GetByKey(businessPlaceId);
            return oBusinessPlace;
        }
    }
}
