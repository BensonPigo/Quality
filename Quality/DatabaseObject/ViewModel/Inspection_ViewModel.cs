using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class Inspection_ViewModel
    {
        public bool Result { get; set; }
        public string ErrMsg { get; set; }
        public string FactoryID { get; set; }
        public string Line { get; set; }
        public DateTime InspectionDate { get; set; }
        public string UserID { get; set; }
        public string StyleID { get; set; }
        public string OrderID { get; set; }
        public string Article { get; set; }
        public string Size { get; set; }
        public string ProductType { get; set; }
        public string ApparelType { get; set; }
        public string ProductTypePMS { get; set; }
        public string Brand { get; set; }
        public string Season { get; set; }
        public string SampleStage { get; set; }
        public string OriginalLine { get; set; }
        public string SizeQty { get; set; }
        public string SizeBalanceQty { get; set; }
        public string OrderQty { get; set; }
        public string OrderBalanceQty { get; set; }
        public string Pass { get; set; }
        public string Reject { get; set; }
        public string Hard { get; set; }
        public string Quick { get; set; }
        public string Wash { get; set; }
        public string Repl { get; set; }
        public string Print { get; set; }
        public string Shade { get; set; }
        public string Dispose { get; set; }
        public List<Top3> top3 { get; set; }
    }

    public class Top3
    {
        public string Defect { get; set; }

        public string Area { get; set; }
    }

    public class Inspection_ChkOrderID_ViewModel
    {
        public string OrderID { get; set; }
        public bool Inpsected { get; set; }
        public bool PulloutComplete { get; set; }
    }
}
