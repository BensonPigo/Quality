using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class WaterAbsorbency_ViewModel : BaseResult
    {
        public WaterAbsorbency_Request Request { get; set; }
        public WaterAbsorbency_Main Main { get; set; }
        public List<WaterAbsorbency_Detail> DetailList { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public string TempFileName { get; set; }
    }

    public class WaterAbsorbency_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
    }

    public class WaterAbsorbency_Main
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
        public string FabricType { get; set; }
        public string FabricDescription { get; set; }
        public string Result { get; set; }
        public string Status { get; set; }

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

        public byte[] TestBeforePicture { get; set; }
        public byte[] TestBeforeWashPicture { get; set; }
        public byte[] TestAfterPicture { get; set; }

        public string MailSubject { get; set; }
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

    public class WaterAbsorbency_Detail : CompareBase
    {
        public long Ukey { get; set; }
        public string ReportNo { get; set; }
        public string EvaluationItem { get; set; }
        public string NoOfDrops { get; set; }
        public decimal Values { get; set; }
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
