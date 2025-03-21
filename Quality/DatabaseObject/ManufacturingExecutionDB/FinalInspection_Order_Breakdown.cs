using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ManufacturingExecutionDB
{
    public class FinalInspection_Order_Breakdown
    {
        public long Ukey { get; set; }
        public string FinalInspectionID { get; set; }
        public string OrderID { get; set; }
        public string Article { get; set; }
        public string SizeCode { get; set; }
        public int LineItem { get; set; }
        public bool Junk { get; set; }
    }
}
