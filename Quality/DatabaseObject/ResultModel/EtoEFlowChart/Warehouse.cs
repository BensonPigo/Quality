using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel.EtoEFlowChart
{
    public class Warehouse
    {
        public string PhysicalInspection { get; set; }
        public List<Warehouse_PhysicalInspection> PhysicalInspection_Detail { get; set; }
    }

    public class Warehouse_PhysicalInspection
    {
        public string OrderID { get; set; }
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Refno { get; set; }
        public string Result { get; set; }
    }
}
