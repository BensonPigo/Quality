using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.RequestModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ResultModel
{
    public class Measurement_ResultModel : Measurement_Request
    {
        public int TotalQty { get; set; }
        public int MeasuredQty { get; set; }
        public int OOTQty { get; set; }
        public string JsonBody { get; set; }
        public List<SelectListItem> Images_Source { get; set; }
        public List<RFT_Inspection_Measurement_Image> Images { get; set; }
    }
    public class RFT_Inspection_Measurement_Image
    {
        public long ID { get; set; }
        public long Seq { get; set; }
        public byte[] Image { get; set; }
        public byte[] TempImage { get; set; }
    }
}
