using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class FIR_Laboratory_Crocking
    {
        /// <summary>ID</summary>
        [Display(Name = "ID")]
        public long ID { get; set; }

        /// <summary>捲號</summary>
        [Display(Name = "捲號")]
        public string Roll { get; set; }

        /// <summary>缸號</summary>
        [Display(Name = "缸號")]
        public string Dyelot { get; set; }

        /// <summary></summary>
        public string DryScale { get; set; }

        /// <summary></summary>
        public string WetScale { get; set; }

        /// <summary>檢驗日期</summary>
        [Display(Name = "檢驗日期")]
        public DateTime? Inspdate { get; set; }

        /// <summary>檢驗人員</summary>
        [Display(Name = "檢驗人員")]
        public string Inspector { get; set; }

        /// <summary>檢驗結果</summary>
        [Display(Name = "檢驗結果")]
        public string Result { get; set; }

        /// <summary>備註</summary>
        [Display(Name = "備註")]
        public string Remark { get; set; }

        /// <summary>新增人員</summary>
        [Display(Name = "新增人員")]
        public string AddName { get; set; }

        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }

        /// <summary>最後編輯人員</summary>
        [Display(Name = "最後編輯人員")]
        public string EditName { get; set; }

        /// <summary>最後編輯時間</summary>
        [Display(Name = "最後編輯時間")]
        public DateTime? EditDate { get; set; }

        /// <summary></summary>
        public string ResultDry { get; set; }

        /// <summary></summary>
        public string ResultWet { get; set; }

        /// <summary></summary>
        public string DryScale_Weft { get; set; }

        /// <summary></summary>
        public string WetScale_Weft { get; set; }

        /// <summary></summary>
        public string ResultDry_Weft { get; set; }

        /// <summary></summary>
        public string ResultWet_Weft { get; set; }

    }
}
