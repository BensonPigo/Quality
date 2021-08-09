using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class DQSReason
    {
        /// <summary>類別</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "類別")]
        public string Type { get; set; }
        /// <summary>ID</summary>
        [Required]
        [StringLength(5)]
        [Display(Name = "ID")]
        public string ID { get; set; }
        /// <summary>閒置原因</summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "閒置原因")]
        public string Description { get; set; }
        /// <summary>閒置原因(當地語言)</summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "閒置原因(當地語言)")]
        public string LocalDescription { get; set; }
        /// <summary>作廢</summary>
        [Required]
        [Display(Name = "作廢")]
        public bool Junk { get; set; }
        /// <summary>建立人員</summary>
        [StringLength(10)]
        [Display(Name = "建立人員")]
        public string AddName { get; set; }
        /// <summary>建立時間</summary>
        [Display(Name = "建立時間")]
        public DateTime? AddDate { get; set; }
        /// <summary>最後編輯人員</summary>
        [StringLength(10)]
        [Display(Name = "最後編輯人員")]
        public string EditName { get; set; }
        /// <summary>最後編輯時間</summary>
        [Display(Name = "最後編輯時間")]
        public DateTime? EditDate { get; set; }

    }
}
