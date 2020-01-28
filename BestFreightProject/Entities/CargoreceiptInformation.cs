using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Entities
{
    public class CargoreceiptInformation
    {
        public string CountryOfOrigin { get; set; }
        public string RegionOfOrigin { get; set; }
        public string CityOfOrigin { get; set; }
        public string Receipt { get; set; }
        public DateTime DepartureDate { get; set; }
    }
}
