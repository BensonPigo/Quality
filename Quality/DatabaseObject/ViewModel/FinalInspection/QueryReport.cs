using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class QueryReport
    {
        public ManufacturingExecutionDB.FinalInspection FinalInspection { get; set; }
        public string SP { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public int TotalSPQty { get; set; }
        public List<FinalInspection_OrderCarton> ListCartonInfo { get; set; }
        public int AvailableQty { get; set; }
        public string AQLPlan  { get; set; }

        public List<FinalInspectionDefectItem> ListDefectItem { get; set; }

        public List<BACriteriaItem> ListBACriteriaItem { get; set; }

        public List<ViewMoistureResult> ListViewMoistureResult { get; set; }

        public string MeasurementUnit { get; set; }

        public List<MeasurementViewItem> ListMeasurementViewItem { get; set; }
    }
}
