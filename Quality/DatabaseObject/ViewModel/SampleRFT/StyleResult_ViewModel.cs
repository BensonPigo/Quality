using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel.SampleRFT
{
    public class StyleResult_ViewModel
    {
        /// <summary>
        /// 以Barcode顯示
        /// </summary>
        public string StyleUkey { get; set; }

        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string ProgramID { get; set; }
        public string Description { get; set; }
        public string Article { get; set; }
        public string StyleName { get; set; }
        public string SpecialMark { get; set; }
        public string SMR { get; set; }
        public string Handle { get; set; }
        public string RFT { get; set; }
        public string ProductType { get; set; }
        public string BuyReadyDate { get; set; }

        public bool Result { get; set; }
        public string MsgScript { get; set; }
        public string StyleRRLRPath { get; set; }
        public string TempFileName { get; set; }

        public bool HasSampleRFTInspection { get; set; }

        public List<StyleResult_SampleRFT> SampleRFT { get; set; }
        public List<StyleResult_FTYDisclaimer> FTYDisclaimer { get; set; }
        public List<StyleResult_RRLR> RRLR { get; set; }
        public List<StyleResult_BulkFGT> BulkFGT { get; set; }
        public List<StyleResult_PoList> PoList { get; set; }
    }

    public class StyleResult_SampleRFT
    {
        public string SP { get; set; }
        public string SampleStage { get; set; }
        public string Factory { get; set; }
        public DateTime? Delivery { get; set; }
        public DateTime? SCIDelivery { get; set; }
        public int InspectedQty { get; set; }
        public string RFT { get; set; }
        public int BAProduct { get; set; }
        public string BAAuditCriteria { get; set; }
    }

    public class StyleResult_FTYDisclaimer
    {
        public string ExpectionFormStatus { get; set; }
        public DateTime? ExpectionFormDate { get; set; }
        public string ExpectionFormRemark { get; set; }
        public string Article { get; set; }
        public string Description { get; set; }
        public string FDFilePath { get; set; }
        public string FDFileName { get; set; }
    }

    public class StyleResult_RRLR
    {
        public string Refno { get; set; }
        public string Supplier { get; set; }
        public string ColorID { get; set; }
        public bool RR { get; set; }
        public string Remark { get; set; }
        public bool LR { get; set; }
    }

    public class StyleResult_BulkFGT
    {
        public string Article { get; set; }
        public string Type { get; set; }
        public string TestName { get; set; }
        public string LastResult { get; set; }
        public DateTime? LastTestDate { get; set; }
    }
    public class StyleResult_PoList
    {
        public string Article { get; set; }
        public string POID { get; set; }
    }
}
