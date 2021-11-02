using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class Others : BaseResult
    {
        public string FinalInspectionID { get; set; }
        public string CFA { get; set; }
        public decimal? ProductionStatus { get; set; }
        public string InspectionResult { get; set; }
        public string ShipmentStatus { get; set; }
        public string OthersRemark { get; set; }

        public List<OtherImage> ListOthersImageItem { get; set; } = new List<OtherImage>();
    }


    public class OtherImage
    {
        public Int64 Ukey { get; set; }
        public string ID { get; set; }
        public byte[] Image { get; set; }

        public byte[] TempImage { get; set; }
        public Int64 RowIndex { get; set; }
    }
}
