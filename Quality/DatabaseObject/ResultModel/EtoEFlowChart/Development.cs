using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel.EtoEFlowChart
{
    public class Development
    {
        public decimal SampleRFT { get; set; }
        public string TestResult { get; set; }
        public string RRLR { get; set; }
        public string FD { get; set; }

        public List<Development_SampleRFT> SampleRFT_Detail { get; set; }
        public List<Development_SampleRFT_CFTComments> CFTComments { get; set; }
        public List<Development_TestResult> TestResult_Detail { get; set; }
        public List<Development_RRLR> RRLR_Detail { get; set; }
        public List<Development_FD> FD_Detail { get; set; }
    }
    public class Development_SampleRFT
    {
        public string OrderID { get; set; }
        public string SampleStage { get; set; }
        public DateTime BuyReadyDate { get; set; }
    }
    public class Development_SampleRFT_CFTComments
    {
        public string OrderID { get; set; }
        public string CommentCategory { get; set; }
        public string Comment { get; set; }
    }
    public class Development_TestResult
    {
        public string Category { get; set; }
        public string ReportNo { get; set; }
        public string Article { get; set; }
        public DateTime TestDate { get; set; }
        public string TestResult { get; set; }
    }
    public class Development_RRLR
    {
        public string Refno { get; set; }
        public string SuppID { get; set; }
        public string ColorID { get; set; }
        public string RR { get; set; }
        public string RRRemark { get; set; }
        public string LR { get; set; }
    }
    public class Development_FD
    {
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string Article { get; set; }
        public string ExpectionFormStatus { get; set; }
        public DateTime ExpectionFormDate { get; set; }
        public string ExpectionFormRemark { get; set; }
    }
}
