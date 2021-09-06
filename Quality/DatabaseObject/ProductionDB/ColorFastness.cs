using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class ColorFastness
    {
        /// <summary>ID</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "ID")]
        public string ID { get; set; }
        /// <summary>採購單號</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "採購單號")]
        public string POID { get; set; }
        /// <summary>檢驗順序</summary>
        [Required]
        [Display(Name = "檢驗順序")]
        public decimal TestNo { get; set; }
        /// <summary>檢驗日期</summary>
        [Required]
        [Display(Name = "檢驗日期")]
        public DateTime? InspDate { get; set; }
        /// <summary>色組</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "色組")]
        public string Article { get; set; }
        /// <summary>檢驗結果</summary>
        [Required]
        [StringLength(15)]
        [Display(Name = "檢驗結果")]
        public string Result { get; set; }
        /// <summary>狀態</summary>
        [StringLength(15)]
        [Display(Name = "狀態")]
        public string Status { get; set; }
        /// <summary>檢驗人員</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "檢驗人員")]
        public string Inspector { get; set; }
        /// <summary>備註</summary>
        [StringLength(120)]
        [Display(Name = "備註")]
        public string Remark { get; set; }
        /// <summary>新增人員</summary>
        [StringLength(10)]
        [Display(Name = "新增人員")]
        public string addName { get; set; }
        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? addDate { get; set; }
        /// <summary>最後編輯人員</summary>
        [StringLength(10)]
        [Display(Name = "最後編輯人員")]
        public string EditName { get; set; }
        /// <summary>最後編輯時間</summary>
        [Display(Name = "最後編輯時間")]
        public DateTime? EditDate { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public int? Temperature { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public int? Cycle { get; set; }
        /// <summary></summary>
        [StringLength(15)]
        [Display(Name = "")]
        public string Detergent { get; set; }
        /// <summary></summary>
        [StringLength(15)]
        [Display(Name = "")]
        public string Machine { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string Drying { get; set; }

        [Display(Name = "測試前的照片")]
        public Byte[] TestBeforePicture { get; set; }

        [Display(Name = "測試後的照片")]
        public Byte[] TestAfterPicture { get; set; }

    }
}
