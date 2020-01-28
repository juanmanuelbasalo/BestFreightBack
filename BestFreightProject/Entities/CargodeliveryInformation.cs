using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Entities
{
    public class CargodeliveryInformation
    {
        public string CountryOfDestination { get; set; }
        public string RegionOfDestination { get; set; }
        public string CityOfDestination { get; set; }
        public string Delivery { get; set; }
        public DateTime ArrivalDate { get; set; }
    }
}
