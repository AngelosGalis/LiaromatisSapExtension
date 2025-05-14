using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LiaromatisSapExtension.Models
{
    public class DeliverySAPPackageModel
    {
        public int DocEntryDev { get; set; }
        public double OpenQty { get; set; }
        public double Quantity { get; set; }
        public string U_AltUoM { get; set; }
        public double U_Quantity { get; set; }
        public double SUMpckQty { get; set; }
        public int DocEntrySO { get; set; }
        public int LineNumSO { get; set; }
    }
}
