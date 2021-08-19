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
        public bool Result { get; set; }
        public string ErrMsg { get; set; }
        public int TotalQty { get; set; }
        public int MeasuredQty { get; set; }
        public int OOTQty { get; set; }
        public List<RFT_Inspection_Measurement> rFT_Inspection_Measurements { get; set; }
    }
}
