using BestFreightProject.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Dtos
{
    public class ExcelCreateDto
    {
        public QuotationInformation QuotationInfo { get; set; }
        public CompanyInformation CompanyInfo { get; set; }
        public CargoInformation CargoInfo { get; set; }
        public CargoreceiptInformation CargoReceiptInfo { get; set; }
        public CargodeliveryInformation CargoDeliveryInfo { get; set; }
        public OceanCarriers OceanCarriersInfo { get; set; }
    }
}
