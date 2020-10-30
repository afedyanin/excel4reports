using System;

namespace SampleApp.Model
{
    public class Order
    {
        public DateTime OrderDate { get; set; }

        public string Region { get; set; }

        public string Rep { get; set; }

        public string Item { get; set; }

        public int Units { get; set; }

        public decimal UnitCost { get; set; }

        public decimal Total { get; set; }
    }
}
