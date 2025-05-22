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
    public class RandomTumblePillingTest_ViewModel : BaseResult
    {
        public RandomTumblePillingTest_Request Request { get; set; }
        public RandomTumblePillingTest_Main Main { get; set; }
        public List<RandomTumblePillingTest_Detail> DetailList { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public List<SelectListItem> TestStandard_Source 
        {
            get
            {
                List<SelectListItem> result = new List<SelectListItem>();
                result.Add(new SelectListItem()
                {
                    Text = "Fleece/Polar Fleece",
                    Value = "Fleece/Polar Fleece"
                });
                result.Add(new SelectListItem()
                {
                    Text = "French Terry",
                    Value = "French Terry"
                });

                return result;
            }
            set
            {

            }
        }
        public List<SelectListItem> Article_Source { get; set; }
        public List<SelectListItem> FaceSideScale_Source
        {
            get
            {
                List<SelectListItem> result = new List<SelectListItem>();
                result.Add(new SelectListItem()
                {
                    Text = "1",
                    Value = "1"
                });
                result.Add(new SelectListItem()
                {
                    Text = "2",
                    Value = "2"
                });
                result.Add(new SelectListItem()
                {
                    Text = "3",
                    Value = "3"
                });
                result.Add(new SelectListItem()
                {
                    Text = "4",
                    Value = "4"
                });
                result.Add(new SelectListItem()
                {
                    Text = "5",
                    Value = "5"
                });

                return result;
            }
            set
            {

            }
        }
        public List<SelectListItem> BackSideScale_Source
        {
            get
            {
                List<SelectListItem> result = new List<SelectListItem>();
                result.Add(new SelectListItem()
                {
                    Text = "1",
                    Value = "1"
                });
                result.Add(new SelectListItem()
                {
                    Text = "1-2",
                    Value = "1-2"
                });
                result.Add(new SelectListItem()
                {
                    Text = "2",
                    Value = "2"
                });
                result.Add(new SelectListItem()
                {
                    Text = "2-3",
                    Value = "2-3"
                });
                result.Add(new SelectListItem()
                {
                    Text = "3",
                    Value = "3"
                });
                result.Add(new SelectListItem()
                {
                    Text = "4",
                    Value = "4"
                });

                return result;
            }
            set
            {

            }
        }

        // 預設摩擦的項目
        public List<string> DefaultSide
        {
            get
            {
                List<string> result = new List<string>()
                {
                    "Face Side","Back Side"
                };
                return result;
            }

        }

        public string TempFileName { get; set; }
    }
    public class RandomTumblePillingTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
    }
    public class RandomTumblePillingTest_Main
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
        
        public string TestStandard { get; set; }
        public string Result { get; set; }
        public string Status { get; set; }
        public string MailSubject { get; set; }

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
        public byte[] TestFaceSideBeforePicture { get; set; }
        public byte[] TestFaceSideAfterPicture { get; set; }
        public byte[] TestBackSideBeforePicture { get; set; }
        public byte[] TestBackSideAfterPicture { get; set; }


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

    public class RandomTumblePillingTest_Detail : CompareBase
    {
        public long Ukey { get; set; }
        public string ReportNo { get; set; }
        public string Side { get; set; }
        public string EvaluationItem { get; set; }
        public string Result { get; set; }
        public string FirstScale { get; set; }
        public string SecondScale { get; set; }
        public bool IsEvenly { get; set; }
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
