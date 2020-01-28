using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Entities
{
    public class OceanCarriers
    {
        private decimal subTotal;
        public List<LogisticService> LogisticServices { get; set; }
        public decimal SubTotal 
        { 
            get => subTotal;
            private set 
            {
                LogisticServices.ForEach(item => subTotal += item.Total);
            } 
        }
        public decimal Taxes { get; set; }
        public decimal Total => SubTotal + (SubTotal * Taxes);
    }
}
