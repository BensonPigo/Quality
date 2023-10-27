using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class WickingHeightTest_ViewModel : BaseResult
    {
        public WickingHeightTest_Request Request { get; set; }
        public WickingHeightTest_Main Main { get; set; }
        public List<WickingHeightTest_Detail> DetailList { get; set; }
        public List<WickingHeightTest_Detail_Item> DetaiItemlList { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public string TempFileName { get; set; }
    }

    public class WickingHeightTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
    }

    public class WickingHeightTest_Main
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
        public string FabricType { get; set; }
        public string FabricRefNo { get; set; }
        public string FabricColor { get; set; }
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


        public byte[] TestWarpPicture { get; set; }
        public byte[] TestWeftPicture { get; set; }


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
    }

    public class WickingHeightTest_Detail : CompareBase
    {
        public long Ukey { get; set; }
        public string ReportNo { get; set; }
        public string EvaluationType { get; set; }
        public decimal WarpAverage { get; set; }
        public string WarpResult { get; set; }
        public decimal WeftAverage { get; set; }
        public string WeftResult { get; set; }
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

    public class WickingHeightTest_Detail_Item : CompareBase
    {
        public long Ukey { get; set; }
        public long WickingHeightTestDetailUkey { get; set; }
        public string EvaluationType { get; set; }
        public string ReportNo { get; set; }
        public string EvaluationItem { get; set; }
        public decimal? WarpValues { get; set; }
        public int? WarpTime { get; set; }
        public decimal? WeftValues { get; set; }
        public int? WeftTime { get; set; }
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
