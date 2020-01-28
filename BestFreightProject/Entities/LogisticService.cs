using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Entities
{
    public class LogisticService
    {
        public string Name { get; set; }
        public int Unit { get; set; }
        public decimal PriceUnit { get; set; }
        public decimal Total => Unit * PriceUnit;
    }
}
