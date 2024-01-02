using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ManufacturingExecutionDB
{
    public class FinalInspection_DefectDetail
    {
        public string FinalInspectionID { get; set; }
        public long ProUkey { get; set; }
        public long DefectCategoryUkey { get; set; }
        public string DefectCategoryDescription { get; set; }
        public int DefectQty { get; set; }
        public string DefectCategoryResult { get; set; }
        public int AcceptedQty { get; set; }
        
    }
}
