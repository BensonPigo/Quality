using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class BACriteria_Result
    {
        public string OrderID { get; set; }
        public string OrderTypeID { get; set; }
        public int Qty { get; set; }
        public int InspectedQty { get; set; }
        public int BAProduct { get; set; }
        public decimal BACriteria { get; set; }

    }
}
