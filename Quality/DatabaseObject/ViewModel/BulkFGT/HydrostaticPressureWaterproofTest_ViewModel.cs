using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class HydrostaticPressureWaterproofTest_ViewModel : BaseResult
    {
        public HydrostaticPressureWaterproofTest_Request Request { get; set; }
        public HydrostaticPressureWaterproofTest_Main Main { get; set; }
        public List<HydrostaticPressureWaterproofTest_Detail> DetailList { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public List<HydrostaticPressureWaterproofStandard> Standard_Source { get; set; }

        /// <summary>
        /// 取得排列好的EvaluationType
        /// </summary>
        /// <returns>排列好的EvaluationType</returns>
        public List<string> EvaluationTypeList
        {
            get
            {
                if (this.Standard_Source != null)
                {
                    return this.Standard_Source.OrderBy(o => o.EvaluationTypeSeq).Select(o => o.EvaluationType).Distinct().ToList();
                }
                else
                {
                    return new List<string>();
                }
            }
        }


        /// <summary>
        /// 取得排列好的EvaluationItem
        /// </summary>
        /// <param name="evaluationType">父層的EvaluationType</param>
        /// <returns>排列好的EvaluationItem</returns>
        public List<string> EvaluationItemList(string evaluationType)
        {
            List<string> result = new List<string>();

            if (this.Standard_Source != null)
            {
                result = this.Standard_Source.Where(o => o.EvaluationType == evaluationType).OrderBy(o => o.EvaluationItemSeq).Select(o => o.EvaluationItem).Distinct().ToList();
            }
            return result;
        }

        public List<SelectListItem> StandardList(string evaluationType, string evaluationItem, string phase)
        {
            List<SelectListItem> result = new List<SelectListItem>();

            if (this.Standard_Source != null)
            {
                var StandardStr = this.Standard_Source
                    .Where(o => o.EvaluationType == evaluationType && o.EvaluationItem == evaluationItem && o.Phase == phase)
                    .Select(o => o.Standard).Distinct().ToList();
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


        public HydrostaticPressureWaterproofStandard GetStandard(string evaluationType, string evaluationItem, string phase)
        {
            HydrostaticPressureWaterproofStandard result = new HydrostaticPressureWaterproofStandard();

            if (this.Standard_Source != null)
            {
                result = this.Standard_Source
                    .Where(o => o.EvaluationType == evaluationType && o.EvaluationItem == evaluationItem && o.Phase == phase && o.Result == "Pass")
                    .FirstOrDefault();
            }
            return result;
        }

        public List<SelectListItem> Article_Source { get; set; }

        public string TempFileName { get; set; }
    }

    public class HydrostaticPressureWaterproofTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
    }
    public class HydrostaticPressureWaterproofTest_Main
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
        public string FabricRefNo { get; set; }
        public string FabricColor { get; set; }
        public string FabricDescription { get; set; }
        public int Temperature { get; set; }
        public string DryingCondition { get; set; }
        public int WashCycles { get; set; }


        public string MailSubject { get; set; }

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
    public class HydrostaticPressureWaterproofTest_Detail : CompareBase
    {
        public long Ukey { get; set; }
        public string ReportNo { get; set; }
        public string EvaluationType { get; set; }
        public string EvaluationItem { get; set; }
        public string AsReceivedValue { get; set; }
        public string AsReceivedResult { get; set; }
        public string AfterWashValue { get; set; }
        public string AfterWashResult { get; set; }
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
        public int EvaluationTypeSeq { get; set; }
        public int EvaluationItemSeq { get; set; }
    }
    public class HydrostaticPressureWaterproofStandard
    {
        public string EvaluationType { get; set; }
        public int EvaluationTypeSeq { get; set; }
        public string EvaluationItem { get; set; }
        public int EvaluationItemSeq { get; set; }
        public string Phase { get; set; }
        public string Standard { get; set; }
        public string Result { get; set; }
    }
}
