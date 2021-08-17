using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class BeautifulProductAudit
    {
        public string FinalInspectionID { get; set; }
        public int? BAQty { get; set; }

        public int? SampleSize { get; set; }
        
        public List<BACriteriaItem> ListBACriteria { get; set; }
    }

    public class BACriteriaItem
    {
        public long Ukey { get; set; }
        public string BACriteria { get; set; }
        public string BACriteriaDesc { get; set; }
        public int? Qty { get; set; }
        public List<byte[]> ListBACriteriaImage { get; set; }
    }
}
