using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Entities
{
    public class CargoInformation
    {
        public int TotalEquipment { get; set; }
        public double Weight { get; set; }
        public string Incoterms { get; set; }
        public string TypeOfCommodity { get; set; }
        public double Cubicfeets { get; set; }
    }
}
