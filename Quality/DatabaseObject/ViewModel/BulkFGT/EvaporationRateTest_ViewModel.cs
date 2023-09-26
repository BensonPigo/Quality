using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class EvaporationRateTest_ViewModel : BaseResult
    {
        public EvaporationRateTest_Request Request { get; set; }
        public EvaporationRateTest_Main Main { get; set; }

        /// <summary>
        /// Before / After
        /// </summary>
        public List<EvaporationRateTest_Detail> DetailList { get; set; }

        /// <summary>
        /// Specimen 1 / Specimen 2 / Specimen 3 .....
        /// </summary>
        public List<EvaporationRateTest_Specimen> SpecimenList { get; set; }


        /// <summary>
        /// Time 0 / 10 / 20 ....
        /// </summary>
        public List<EvaporationRateTest_Specimen_Time> TimeList { get; set; }

        /// <summary>
        /// R1 / R2 / R3 .....
        /// </summary>
        public List<EvaporationRateTest_Specimen_Rate> RateList { get; set; }

        public List<SelectListItem> ReportNo_Source { get; set; }
        public List<SelectListItem> Article_Source { get; set; }
        public string TempFileName { get; set; }
    }

    public class EvaporationRateTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
    }
    public class EvaporationRateTest_Main
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
        public decimal BeforeAverageRate { get; set; }
        public decimal AfterAverageRate { get; set; }


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
    public class EvaporationRateTest_Detail : CompareBase
    {
        public string ReportNo { get; set; }
        public string Type { get; set; }
        public long Ukey { get; set; }
        public decimal EvaporationRateAverage { get; set; }
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
    public class EvaporationRateTest_Specimen : CompareBase
    {
        /// <summary>
        /// 用於前端對應父層
        /// </summary>
        public string DetailType { get; set; }
        public long DetailUkey { get; set; }
        public string SpecimenID { get; set; }
        public long Ukey { get; set; }
        public decimal RateAverage { get; set; }
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
    public class EvaporationRateTest_Specimen_Time : CompareBase
    {
        public string DetailType { get; set; }
        /// <summary>
        /// 用於前端對應父層
        /// </summary>
        public string SpecimenID{ get; set; }
        public long SpecimenUkey { get; set; }
        public int Time { get; set; }
        public bool IsInitialMass { get; set; }
        public long Ukey { get; set; }
        public decimal Mass { get; set; }
        public decimal Evaporation { get; set; }
        public long InitialTimeUkey { get; set; }

        /// <summary>
        /// 新增資料時，都還沒有Ukey，因此無法找到初始值的Ukey，只好透過 Time去尋找，因為同一個Specimen只會有一個Time
        /// </summary>
        public int InitialTime { get; set; }

        /// <summary>
        /// Time 0 ~ 40 是基礎資料，不可被刪除
        /// </summary>
        public bool CanDelete
        {
            get
            {
                if (this.Time > 40)
                {
                    return true;
                }
                else
                {
                    return false;
                }
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
    }
    public class EvaporationRateTest_Specimen_Rate : CompareBase
    {
        public string DetailType { get; set; }
        /// <summary>
        /// 用於前端對應父層
        /// </summary>
        public string SpecimenID { get; set; }
        public long SpecimenUkey { get; set; }
        public string RateName { get; set; }
        public long Ukey { get; set; }
        public decimal Value { get; set; }
        public long Subtrahend_TimeUkey { get; set; }
        public long Minuend_TimeUkey { get; set; }

        /// <summary>
        /// 新增資料時，都還沒有Ukey，因此無法透過Ukey去找減數，只好透過 Time去尋找，因為同一個Specimen只會有一個Time
        /// </summary>
        public int Subtrahend_Time { get; set; }

        /// <summary>
        /// 新增資料時，都還沒有Ukey，因此無法透過Ukey去找被減數，只好透過 Time去尋找，因為同一個Specimen只會有一個Time
        /// </summary>
        public int Minuend_Time { get; set; }
        public int Ratio { get; set; }
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
