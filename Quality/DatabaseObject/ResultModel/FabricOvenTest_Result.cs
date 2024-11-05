using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class FabricOvenTest_Result : BaseResult
    {
        public FabricOvenTest_Main Main { get; set; }
        public List<FabricOvenTest_Detail> Details { get; set; } = new List<FabricOvenTest_Detail>();
    }

    public class FabricOvenTest_Main
    {
        public string POID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public DateTime? CutInline { get; set; }
        public DateTime? MinSciDelivery { get; set; }
        public DateTime? TargetLeadTime { get; set; }
        public DateTime? CompletionDate { get; set; }
        public decimal LabOvenPercent { get; set; }
        public string Remark { get; set; }
        public string CreateBy { get; set; }
        public string EditBy { get; set; }
    }

    public class FabricOvenTest_Detail
    {
        public string TestNo { get; set; }
        public string ReportNo { get; set; }
        public DateTime? InspDate { get; set; }
        public string Article { get; set; }
        public string Result { get; set; }
        public string Inspector { get; set; }
        public string Remark { get; set; }
        public string LastUpdate { get; set; }
        public string Status { get; set; }
    }

    public class FabricOvenTest_Detail_Result : BaseResult
    {
        public string MDivisionID { get; set; }
        public List<string> ScaleIDs { get; set; }
        public FabricOvenTest_Detail_Main Main { get; set; } = new FabricOvenTest_Detail_Main();
        public List<FabricOvenTest_Detail_Detail> Details { get; set; } = new List<FabricOvenTest_Detail_Detail>();
    }

    public class FabricOvenTest_Detail_Main
    {
        public string TestNo { get; set; }
        public string POID { get; set; }
        public string BrandID { get; set; }
        public string ReportNo { get; set; }
        public DateTime? InspDate { get; set; }
        public string Article { get; set; }
        public string Inspector { get; set; }
        public string InspectorName { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string Status { get; set; }
        public byte[] TestBeforePicture { get; set; }
        public byte[] TestAfterPicture { get; set; }
        public string MailSubject { get; set; }
        public string Approver { get; set; }
        public string ApproverName { get; set; }
        public DateTime? ReportDate { get; set; }
    }

    public class FabricOvenTest_Detail_Detail : CompareBase
    {
        public DateTime? SubmitDate { get; set; }
        public string OvenGroup { get; set; }
        public string SEQ { get; set; }
        public string Roll { get; set; }
        public string Dyelot { get; set; }
        public string Refno { get; set; }
        public string SCIRefno { get; set; }
        public string ColorID { get; set; }
        public string Result { get; set; }
        public string ChangeScale { get; set; }
        public string ResultChange { get; set; }
        public string StainingScale { get; set; }
        public string ResultStain { get; set; }
        public string Remark { get; set; }
        public string LastUpdate { get; set; }
        public int Temperature { get; set; }
        public int Time { get; set; }

        public string Seq1 {
            get {
                if (string.IsNullOrEmpty(this.SEQ))
                {
                    return string.Empty;
                }

                if (!this.SEQ.Contains("-"))
                {
                    return string.Empty;
                }

                return this.SEQ.Split('-')[0].Trim();
            }
        }

        public string Seq2
        {
            get
            {
                if (string.IsNullOrEmpty(this.SEQ))
                {
                    return string.Empty;
                }

                if (!this.SEQ.Contains("-"))
                {
                    return string.Empty;
                }

                return this.SEQ.Split('-')[1].Trim();
            }
        }
    }
}
