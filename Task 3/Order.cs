using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task_3
{
    internal class Order
    {
        public string OrderCode { get; set; }
        public string ProductCode { get; set; }
        public string ClientCode { get; set; }
        public string RequestNumber { get; set; }
        public string Quantity { get; set; }
        public DateTime Date { get; set; }
    }
}
