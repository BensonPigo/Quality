using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class Setting : BaseResult
    {
        public string FinalInspectionID { get; set; } = string.Empty;
        public string InspectionStage { get; set; } = string.Empty;
        public DateTime? AuditDate { get; set; } 
        public string SewingLineID { get; set; } = string.Empty;
        public string InspectionTimes { get; set; } = string.Empty;

        public string AcceptableQualityLevelsUkey { get; set; } = string.Empty;
        public int? SampleSize { get; set; }
        public int? AcceptQty { get; set; }

        public List<SelectedPO> SelectedPO { get; set; }

        public List<SelectCarton> SelectCarton { get; set; }

        public List<AcceptableQualityLevels> AcceptableQualityLevels { get; set; }
    }

    public class SelectedPO
    {
        public string OrderID { get; set; }
        public string POID { get; set; }
        public string StyleID { get; set; }
        public string SeasonID { get; set; }
        public string BrandID { get; set; }
        public int Qty { get; set; }
        public string Cartons { get; set; }
        public int AvailableQty { get; set; }
    }

    public class SelectCarton
    {
        public bool Selected { get; set; }
        public string OrderID { get; set; }
        public string PackingListID { get; set; }
        public string CTNNo { get; set; }
    }

}
