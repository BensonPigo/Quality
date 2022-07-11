using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class QueryFinalInspection_ViewModel : ResultModelBase<QueryFinalInspection>
    {
        public string SP { get; set; }
        public string CustPONO { get; set; }
        public string StyleID { get; set; }
        public string InspectionResult { get; set; }        
        public DateTime? AuditDateStart { get; set; }
        public DateTime? AuditDateEnd { get; set; }
        public bool ExcludeJunk { get; set; }
    }

    public class QueryFinalInspection
    {
        public string FinalInspectionID { get; set; }
        public string SP { get; set; }
        public string CustPONO { get; set; }
        public string AuditDate { get; set; }
        public string SPQty { get; set; }
        public string StyleID { get; set; }
        public string Season { get; set; }
        public string BrandID { get; set; }
        public string Article { get; set; }
        public string InspectionTimes { get; set; }
        public string InspectionStage { get; set; }
        public string InspectionResult { get; set; }
        public string IsTransferToPMS { get; set; }
        public string IsTransferToPivot88 { get; set; }
    }
}
