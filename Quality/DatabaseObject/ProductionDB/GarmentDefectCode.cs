using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class GarmentDefectCode
    {
        /// <summary>瑕疵代號</summary>
        [Required]
        [StringLength(3)]
        [Display(Name = "瑕疵代號")]
        public string ID { get; set; }
        /// <summary>描述</summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "描述")]
        public string Description { get; set; }
        /// <summary>瑕疵種類</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "瑕疵種類")]
        public string GarmentDefectTypeID { get; set; }
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
        
        public bool Junk { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(100)]
        
        public string LocalDescription { get; set; }
        /// <summary></summary>
        [Required]
        
        public Byte Seq { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        public string ReworkTotalFailCode { get; set; }
        /// <summary></summary>
        [Required]
        public bool IsCFA { get; set; }

    }
}
