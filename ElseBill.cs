using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace chatupon
{
    class ElseBill
    {
        public int Number { get; set; }
        public int Quantity { get; set; }
        public string Descriiption { get; set; }
        public double UnitPrice { get; set; }
        public double Amount { get; set; }

        internal static void RemoveAll(Func<object, bool> p)
        {
            throw new NotImplementedException();
        }
    }
}
