using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class Measurement_ResultModel : Measurement_Request
    {
        public int TotalQty { get; set; }
        public int MeasuredQty { get; set; }
        public int OOTQty { get; set; }
        public string JsonBody { get; set; }
    }
}
