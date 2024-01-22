using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class QueryReport : BaseResult
    {
        public ManufacturingExecutionDB.FinalInspection FinalInspection { get; set; }
        public string SP { get; set; }
        public string StyleID { get; set; }
        public string SeasonID { get; set; }
        public string BrandID { get; set; }
        public string Dest { get; set; }
        public int TotalSPQty { get; set; }
        public List<FinalInspection_OrderCarton> ListCartonInfo { get; set; }
        public int AvailableQty { get; set; }
        public string AQLPlan  { get; set; }


        public string GoOnInspectURL { get; set; }

        public List<FinalInspectionDefectItem> ListDefectItem { get; set; }

        public List<FinalInspection_DefectDetail> FinalInspection_DefectDetails { get; set; }

        public List<BACriteriaItem> ListBACriteriaItem { get; set; }

        public List<ViewMoistureResult> ListViewMoistureResult { get; set; }

        public string MeasurementUnit { get; set; }

        public List<MeasurementViewItem> ListMeasurementViewItem { get; set; }

        public List<SelectOrderShipSeq> ListShipModeSeq { get; set; }
        public List<FinalInspectionSignature> ListFinalInspectionSignature{ get; set; }

        public string GetFinalInspectionSignatureByJobTitle(string JobTitle)
        {
            List<string> rtn = new List<string>();
            if (this.ListFinalInspectionSignature.Any())
            {
                rtn = this.ListFinalInspectionSignature.Where(o => o.JobTitle == JobTitle).Select(o => o.UserID).Distinct().ToList();
            }

            string result = string.Join(",", rtn);

            return result;
        }
        public string SignatureBy_QC { get; set; }
        public string SignatureBy_QCManager { get; set; }
        public string SignatureBy_Production { get; set; }
        public string SignatureBy_ProductionManager { get; set; }
        public string SignatureBy_TSD { get; set; }

        public List<SelectListItem> JobTitleList { get; set; }
    }
}
