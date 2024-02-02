using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class BeautifulProductAudit : BaseResult
    {
        public string FinalInspectionID { get; set; }

        public string InspectionStep { get; set; }
        public int BAQty { get; set; }

        public int SampleSize { get; set; }
        
        public List<BACriteriaItem> ListBACriteria { get; set; }
    }

    public class BACriteriaItem
    {
        public long Ukey { get; set; }
        public string BACriteria { get; set; }
        public string BACriteriaDesc { get; set; }
        public int Qty { get; set; }
        public bool HasImage { get; set; }
        //public List<byte[]> ListBACriteriaImage { get; set; } = new List<byte[]>();
        public List<ImageRemark> ListBACriteriaImage { get; set; } = new List<ImageRemark>();
        public byte[] TempImage { get; set; }
        public string TempRemark { get; set; }
        public long RowIndex { get; set; }
        public string LoginToken { get; set; }
    }
    //public class ImageRemark
    //{
    //    public byte[] Image { get; set; }
    //    public string Remark { get; set; }
    //}
}
