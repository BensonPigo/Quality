using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class AddDefect : BaseResult
    {
        public string FinalInspectionID { get; set; }
        public int? RejectQty { get; set; }
        public string InspectionStep { get; set; } = "";

        public int? SampleSize { get; set; }

        public List<FinalInspectionDefectItem> ListFinalInspectionDefectItem { get; set; }

    }

    public class FinalInspectionDefectItem
    {
        public long Ukey { get; set; }
        public string DefectType { get; set; }

        public string DefectCode { get; set; }

        public string DefectTypeDesc { get; set; }

        public string DefectCodeDesc { get; set; }

        public int Qty { get; set; }
        public bool HasImage { get; set; }

        //public List<byte[]> ListFinalInspectionDefectImage { get; set; } = new List<byte[]>();
        public List<ImageRemark> ListFinalInspectionDefectImage { get; set; } = new List<ImageRemark>();

        public byte[] TempImage { get; set; }
        public string TempRemark { get; set; }
        public Int64 RowIndex { get; set; }
        public string LoginToken { get; set; }
    }
    public class ImageRemark
    {
        public byte[] Image { get; set; }
        public string Remark { get; set; }
    }
}
