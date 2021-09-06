using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{

    public class ColorFastness_Detail
    {
        /// <summary>ID</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "ID")]
        public string ID { get; set; }

        /// <summary>分組</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "分組")]
        public string ColorFastnessGroup { get; set; }

        /// <summary>大小項</summary>
        [Required]
        [StringLength(3)]
        [Display(Name = "大小項")]
        public string SEQ1 { get; set; }

        /// <summary></summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "")]
        public string SEQ2 { get; set; }

        /// <summary>捲號</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "捲號")]
        public string Roll { get; set; }

        /// <summary>缸號</summary>
        [StringLength(8)]
        [Display(Name = "缸號")]
        public string Dyelot { get; set; }

        /// <summary>結果</summary>
        [StringLength(4)]
        [Display(Name = "結果")]
        public string Result { get; set; }

        /// <summary>色差灰階</summary>
        [StringLength(8)]
        [Display(Name = "色差灰階")]
        public string changeScale { get; set; }

        /// <summary>染色灰階</summary>
        [StringLength(8)]
        [Display(Name = "染色灰階")]
        public string StainingScale { get; set; }

        /// <summary>備註</summary>
        [StringLength(60)]
        [Display(Name = "備註")]
        public string Remark { get; set; }

        /// <summary>新增者</summary>
        [StringLength(10)]
        [Display(Name = "新增者")]
        public string AddName { get; set; }

        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }

        /// <summary>編輯者</summary>
        [StringLength(10)]
        [Display(Name = "編輯者")]
        public string EditName { get; set; }

        /// <summary>編輯時間</summary>
        [Display(Name = "編輯時間")]
        public DateTime? EditDate { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? SubmitDate { get; set; }

        /// <summary></summary>
        [StringLength(5)]
        [Display(Name = "")]
        public string ResultChange { get; set; }

        /// <summary></summary>
        [StringLength(5)]
        [Display(Name = "")]
        public string ResultStain { get; set; }


    }
}
