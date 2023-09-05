using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class StickerTest_ViewModel : BaseResult
    {
        public StickerTest_Request Request { get; set; }
        public StickerTest_Main Main { get; set; }
        public List<StickerTest_Detail> DetailList { get; set; }
        public List<StickerTest_Detail_Item> DetailItemList { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public List<SelectListItem> Scale_Source { get; set; }

        public List<SelectListItem> TestStandard_Source
        {
            get
            {
                List<SelectListItem> result = new List<SelectListItem>();
                result.Add(new SelectListItem()
                {
                    Text = "Residue",
                    Value = "Residue"
                });
                result.Add(new SelectListItem()
                {
                    Text = "Aging",
                    Value = "Aging"
                });

                return result;
            }
            set
            {

            }
        }

        public List<StickerTestItem> Item_Source { get; set; }

        /// <summary>
        /// 取得排列好的EvaluationItem
        /// </summary>
        /// <param name="evaluationType">父層的EvaluationType</param>
        /// <returns>排列好的EvaluationItem</returns>
        public List<string> EvaluationItemList()
        {
            List<string> result = new List<string>();

            if (this.Item_Source != null)
            {
                result = this.Item_Source.OrderBy(o => o.EvaluationItemSeq).Select(o => o.EvaluationItem).Distinct().ToList();
            }
            return result;
        }

        public List<SelectListItem> ItemList(string evaluationItem, string evaluationItemDesc)
        {
            List<SelectListItem> result = new List<SelectListItem>();

            if (this.Item_Source != null)
            {
                var StandardStr = this.Item_Source
                    .Where(o => o.EvaluationItem == evaluationItem && o.EvaluationItemDesc == evaluationItemDesc)
                    .Select(o => o.Value).Distinct().ToList();
                foreach (var standard in StandardStr)
                {
                    result.Add(new SelectListItem()
                    {
                        Text = standard,
                        Value = standard
                    });
                }
            }
            return result;
        }


        public StickerTestItem GetStandard(string evaluationItem, string evaluationItemDesc)
        {
            StickerTestItem result = new StickerTestItem();

            if (this.Item_Source != null)
            {
                result = this.Item_Source
                    .Where(o => o.EvaluationItem == evaluationItem && o.EvaluationItemDesc == evaluationItemDesc && o.Result == "Pass")
                    .FirstOrDefault();
            }
            return result;
        }

        public List<SelectListItem> Article_Source { get; set; }

        public string TempFileName { get; set; }
    }

    public class StickerTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
    }
    public class StickerTest_Main
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
                if (string.IsNullOrEmpty(this.Seq1) || string.IsNullOrWhiteSpace(this.Seq2))
                {
                    return string.Empty;
                }
                else
                {
                    return $@"{this.Seq1}-{this.Seq2}"; ;
                }
            }

        }
        public string TestStandard { get; set; }
        public string FabricRefNo { get; set; }
        public string FabricColor { get; set; }
        public string FabricDescription { get; set; }
        public string AccRefNo { get; set; }
        public string AccColor { get; set; }
        public int Temperature { get; set; }
        public int Time { get; set; }
        public decimal Humidity { get; set; }



        public string Remark { get; set; }
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
    public class StickerTest_Detail : CompareBase
    {
        public string ReportNo { get; set; }
        public string EvaluationItem { get; set; }
        public string Scale { get; set; }
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

    public class StickerTest_Detail_Item : CompareBase
    {
        public string ReportNo { get; set; }
        public string EvaluationItem { get; set; }
        public string ItemID { get; set; }
        public string EvaluationItemDesc { get; set; }
        public string Value { get; set; }
        public string Result { get; set; }
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


    public class StickerTestItem
    {
        public int EvaluationItemSeq { get; set; }
        public string EvaluationItem { get; set; }
        public int EvaluationItemDescSeq { get; set; }
        public string EvaluationItemDesc { get; set; }
        public string Value { get; set; }
        public string Result { get; set; }
    }
}
