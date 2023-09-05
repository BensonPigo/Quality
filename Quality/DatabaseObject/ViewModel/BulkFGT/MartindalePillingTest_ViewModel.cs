using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class MartindalePillingTest_ViewModel : BaseResult
    {
        public MartindalePillingTest_Request Request { get; set; }
        public MartindalePillingTest_Main Main { get; set; }
        public List<MartindalePillingTest_Detail> DetailList { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public List<SelectListItem> TestStandard_Source 
        {
            get
            {
                List<SelectListItem> result = new List<SelectListItem>();
                if (this.Main != null && this.Main.FabricType == "WOVEN")
                {
                    result.Add(new SelectListItem()
                    {
                        Text="W1",Value="W1"
                    });
                    result.Add(new SelectListItem()
                    {
                        Text = "W2",
                        Value = "W2"
                    });
                    result.Add(new SelectListItem()
                    {
                        Text = "W3",
                        Value = "W3"
                    });
                }
                else if (this.Main != null && this.Main.FabricType == "KNIT")
                {
                    result.Add(new SelectListItem()
                    {
                        Text = "K1",
                        Value = "K1"
                    });
                    result.Add(new SelectListItem()
                    {
                        Text = "K2",
                        Value = "K2"
                    });
                    result.Add(new SelectListItem()
                    {
                        Text = "K3",
                        Value = "K3"
                    });
                }

                return result;
            }
            set
            {

            }
        }
        public List<SelectListItem> Article_Source { get; set; }
        public List<SelectListItem> Scale_Source { get; set; }

        // 預設摩擦的項目
        public List<string> DefaultRubTimes 
        {
            get
            {
                List<string> result = new List<string>()
                {
                    "500 Rub","2000 Rub"
                };
                return result;
            }

        }

        public string TempFileName { get; set; }
    }
    public class MartindalePillingTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
    }
    public class MartindalePillingTest_Main
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
        
        // 分W1,W2,W3,K1,K2,K3
        public string TestStandard { get; set; }
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
        public byte[] Test500AfterPicture { get; set; }
        public byte[] Test2000AfterPicture { get; set; }


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

    public class MartindalePillingTest_Detail : CompareBase
    {
        public long Ukey { get; set; }
        public string ReportNo { get; set; }
        public string EvaluationItem { get; set; }
        public string RubTimes { get; set; }
        public string Result { get; set; }
        public string Scale { get; set; }
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
