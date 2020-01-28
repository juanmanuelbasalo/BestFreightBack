using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Entities
{
    public class QuotationInformation
    {
        public string FreightType { get; set; }
        public string SubFreightType { get; set; }
        public int QuotationNumber { get; set; }
        public DateTime QuotationDate { get; set; }
        public string SpecialInstructions { get; set; }
    }
}
