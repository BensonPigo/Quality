using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class Others
    {
        public string FinalInspectionID { get; set; }
        public string CFA { get; set; }
        public string ProductionStatus { get; set; }
        public string InspectionResult { get; set; }
        public string ShipmentStatus { get; set; }
        public string OthersRemark { get; set; }
        
        public List<byte[]> ListOthersImageItem { get; set; }
    }
}
