using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class QueryFinalInspection
    {
        public string FinalInspectionID { get; set; }
        public string SP { get; set; }
        public string POID { get; set; }
        public string SPQty { get; set; }
        public string StyleID { get; set; }
        public string Season { get; set; }
        public string BrandID { get; set; }
        public string InspectionTimes { get; set; }
        public string InspectionStage { get; set; }
        public string InspectionResult { get; set; }
    }
}
