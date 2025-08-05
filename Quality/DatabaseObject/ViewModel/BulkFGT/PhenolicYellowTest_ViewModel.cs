using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class PhenolicYellowTest_ViewModel : BaseResult
    {
        public PhenolicYellowTest_Request Request { get; set; }
        public PhenolicYellowTest_Main Main { get; set; }
        public List<PhenolicYellowTest_Detail> DetailList { get; set; }
        public List<SelectListItem> Temperature_Source { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public List<SelectListItem> Scale_Source { get; set; }
        public List<SelectListItem> Article_Source { get; set; }


        public string TempFileName { get; set; }
    }
    public class PhenolicYellowTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
    }
    public class PhenolicYellowTest_Main
    {
        public string ReportNo { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string FactoryID { get; set; }
        public string OrderID { get; set; }
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Seq 
        {
            get
            {
                return $@"{this.Seq1}-{this.Seq2}"; ;
            }

        }
        public string FabricRefNo { get; set; }
        public string FabricColor { get; set; }
        public string Status { get; set; }
        public string Result { get; set; }
        public decimal Temperature { get; set; }
        public decimal Time { get; set; }
        public decimal Humidity { get; set; }

        /// <summary>
        /// User Encode日期，系統寫入
        /// </summary>
        public DateTime? ReportDate { get; set; }
        public string ReportDateText
        {
            get
            {
                string rtn = string.Empty;
                if (this.ReportDate.HasValue)
                {
                    rtn = this.ReportDate.Value.ToString("yyyy/MM/dd");
                }
                return rtn;
            }

        }

        /// <summary>
        /// User 自行寫入
        /// </summary>
        public DateTime? SubmitDate { get; set; }
        public string SubmitDateText
        {
            get
            {
                string rtn = string.Empty;
                if (this.SubmitDate.HasValue)
                {
                    rtn = this.SubmitDate.Value.ToString("yyyy/MM/dd");
                }
                return rtn;
            }

        }

        public string MailSubject { get; set; }

        public byte[] TestBeforePicture { get; set; }
        public byte[] TestAfterPicture { get; set; }


        public string AddName { get; set; }
        public DateTime? AddDate { get; set; }
        public string CreateBy
        {
            get
            {
                return this.AddDate.HasValue ? $"{this.AddDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.AddName}" : string.Empty;
            }
        }

        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastUpadate
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.EditName}" : string.Empty;
            }
        }

        public string EditType { get; set; }
        public string Approver { get; set; }
        public string ApproverName { get; set; }
        public string Preparer { get; set; }
        public string PreparerName { get; set; }
    }

    public class PhenolicYellowTest_Detail : CompareBase
    {
        public long Ukey { get; set; }
        public string ReportNo { get; set; }
        public string EvaluationItem { get; set; }
        public string Roll { get; set; }
        public string Dyelot { get; set; }
        public string Scale { get; set; }
        public string Status { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastUpadate
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.EditName}" : string.Empty;
            }
        }
    }

}
