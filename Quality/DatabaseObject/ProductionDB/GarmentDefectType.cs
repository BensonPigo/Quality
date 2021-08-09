using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class GarmentDefectType
    {
        /// <summary>瑕疵代號</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "瑕疵代號")]
        public string ID { get; set; }
        /// <summary>描述</summary>
        [Required]
        [StringLength(60)]
        [Display(Name = "描述")]
        public string Description { get; set; }
        /// <summary>取消</summary>
        [Required]
        [Display(Name = "取消")]
        public bool Junk { get; set; }
        /// <summary>新增人員</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "新增人員")]
        public string AddName { get; set; }
        /// <summary>新增時間</summary>
        [Required]
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }
        /// <summary>最後修改人員</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "最後修改人員")]
        public string EditName { get; set; }
        /// <summary>最後修改時間</summary>
        [Display(Name = "最後修改時間")]
        public DateTime? EditDate { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "")]
        public string LocalDescription { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public Byte Seq { get; set; }

    }
}
