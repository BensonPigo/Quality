using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class RFT_Inspection
    {
        /// <summary>ID</summary>
        [Required]
        [Display(Name = "ID")]
        public long ID { get; set; }
        /// <summary>訂單號碼</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "訂單號碼")]
        public string OrderID { get; set; }
        /// <summary>色組</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "色組")]
        public string Article { get; set; }
        /// <summary>部位 (T, B, I, O)</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "部位 (T, B, I, O)")]
        public string Location { get; set; }
        /// <summary>尺碼</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "尺碼")]
        public string Size { get; set; }
        /// <summary>產線</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "產線")]
        public string Line { get; set; }
        /// <summary>廠代</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "廠代")]
        public string FactoryID { get; set; }
        /// <summary>款式</summary>
        [Required]
        [Display(Name = "款式")]
        public long StyleUkey { get; set; }

        /// <summary>款式名稱</summary>
        [Required]
        [Display(Name = "款式名稱")]
        public string Style { get; set; }
        /// <summary>返修狀態 (Hard, Quick, Print, Quick, Repl., Shade, Wash)</summary>
        [Required]
        [StringLength(5)]
        [Display(Name = "返修狀態 (Hard, Quick, Print, Quick, Repl., Shade, Wash)")]
        public string FixType { get; set; }
        /// <summary>返修卡號</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "返修卡號")]
        public string ReworkCardNo { get; set; }
        /// <summary>檢驗狀態 (Pass, Reject, Fixed, Delete, Dispose)</summary>
        [Required]
        [StringLength(7)]
        [Display(Name = "檢驗狀態 (Pass, Reject, Fixed, Delete, Dispose)")]
        public string Status { get; set; }
        /// <summary>新增日期</summary>
        [Display(Name = "新增日期")]
        public DateTime? AddDate { get; set; }
        /// <summary>新增人員</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "新增人員")]
        public string AddName { get; set; }
        /// <summary>編輯日期</summary>
        [Display(Name = "編輯日期")]
        public DateTime? EditDate { get; set; }
        /// <summary>編輯人員</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "編輯人員")]
        public string EditName { get; set; }
        /// <summary>返修卡狀態 (Hard, Quick)</summary>
        [Required]
        [StringLength(5)]
        [Display(Name = "返修卡狀態 (Hard, Quick)")]
        public string ReworkCardType { get; set; }
        /// <summary>實際產出日</summary>
        [Required]
        [Display(Name = "實際產出日")]
        public DateTime? InspectionDate { get; set; }
        /// <summary>報廢原因</summary>
        [Required]
        [StringLength(5)]
        [Display(Name = "報廢原因")]
        public string DisposeReason { get; set; }

    }
}
