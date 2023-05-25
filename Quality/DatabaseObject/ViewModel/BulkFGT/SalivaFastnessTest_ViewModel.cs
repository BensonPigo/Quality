using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class SalivaFastnessTest_ViewModel : BaseResult
    {
        public SalivaFastnessTest_Request Request { get; set; }
        public SalivaFastnessTest_Main Main { get; set; }
        public List<SalivaFastnessTest_Detail> DetailList { get; set; }
        public List<SelectListItem> Temperature_Source { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public List<SelectListItem> Scale_Source { get; set; }
        public List<SelectListItem> ItemTested_Source
        {
            get
            {
                return new List<SelectListItem> {
                    new SelectListItem(){Text="Fabric",Value="Fabric"},
                    new SelectListItem(){Text="Accessory",Value="Accessory"},
                    new SelectListItem(){Text="Printing",Value="Printing"},
                };
            }

        }
        public List<SelectListItem> Article_Source { get; set; }


        public string TempFileName { get; set; }
    }
    public class SalivaFastnessTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
    }
    public class SalivaFastnessTest_Main
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
        public string ItemTested { get; set; }
        public string FabricRefNo { get; set; }
        public string FabricColor { get; set; }
        public string FabricDescription { get; set; }
        public string TypeOfPrint { get; set; }
        public string PrintColor { get; set; }
        public decimal Temperature { get; set; }
        public decimal Time { get; set; }
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
    }

    public class SalivaFastnessTest_Detail : CompareBase
    {
        public long Ukey { get; set; }
        public string ReportNo { get; set; }
        public string EvaluationItem { get; set; }
        public string AllResult { get; set; }

        public string AcetateScale { get; set; }
        public string CottonScale { get; set; }
        public string NylonScale { get; set; }
        public string PolyesterScale { get; set; }
        public string AcrylicScale { get; set; }
        public string WoolScale { get; set; }

        public string AcetateResult { get; set; }
        public string CottonResult { get; set; }
        public string NylonResult { get; set; }
        public string PolyesterResult { get; set; }
        public string AcrylicResult { get; set; }
        public string WoolResult { get; set; }
        public string Status { get; set; }
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
