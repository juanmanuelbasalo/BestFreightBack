using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Entities
{
    public class Excel
    {
        public QuotationInformation QuotationInfo { get; set; }
        public CompanyInformation CompanyInfo { get; set; }
        public CargoInformation CargoInfo { get; set; }
        public CargoreceiptInformation CargoReceiptInfo { get; set; }
        public CargodeliveryInformation CargoDeliveryInfo { get; set; }
        public List<OceanCarriers> OceanCarriersInfo { get; set; }
    }
}
