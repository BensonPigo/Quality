using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Oven
    {
        /// <summary>ID</summary>
        [Display(Name = "ID")]
        public Int64? ID { get; set; }

        /// <summary>採購單號</summary>
        [Display(Name = "採購單號")]
        public string POID { get; set; }

        /// <summary>檢驗順序</summary>
        [Display(Name = "檢驗順序")]
        public decimal TestNo { get; set; }

        /// <summary>檢驗日期</summary>
        [Display(Name = "檢驗日期")]
        public DateTime? InspDate { get; set; }

        /// <summary>色組</summary>
        [Display(Name = "色組")]
        public string Article { get; set; }

        /// <summary>檢驗結果</summary>
        [Display(Name = "檢驗結果")]
        public string Result { get; set; }

        /// <summary>狀態</summary>
        [Display(Name = "狀態")]
        public string Status { get; set; }

        /// <summary>檢驗人員</summary>
        [Display(Name = "檢驗人員")]
        public string Inspector { get; set; }

        /// <summary>備註</summary>
        [Display(Name = "備註")]
        public string Remark { get; set; }

        /// <summary>新增人員</summary>
        [Display(Name = "新增人員")]
        public string addName { get; set; }

        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? addDate { get; set; }

        /// <summary>最後編輯人員</summary>
        [Display(Name = "最後編輯人員")]
        public string EditName { get; set; }

        /// <summary>最後編輯時間</summary>
        [Display(Name = "最後編輯時間")]
        public DateTime? EditDate { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public int? Temperature { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public int? Time { get; set; }

        /// <summary>測試前的照片</summary>
        [Display(Name = "測試前的照片")]
        public Byte[] TestBeforePicture { get; set; }

        /// <summary>測試後的照片</summary>
        [Display(Name = "測試後的照片")]
        public Byte[] TestAfterPicture { get; set; }

    }
}
