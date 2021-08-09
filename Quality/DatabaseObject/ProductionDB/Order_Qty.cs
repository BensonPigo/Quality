using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Order_Qty
    {
        /// <summary>訂單編號</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "訂單編號")]
        public string ID { get; set; }
        /// <summary>顏色組</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "顏色組")]
        public string Article { get; set; }
        /// <summary>尺寸</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "尺寸")]
        public string SizeCode { get; set; }
        /// <summary>數量</summary>
        [Required]
        [Display(Name = "數量")]
        public int? Qty { get; set; }
        /// <summary>原始數量</summary>
        [Required]
        [Display(Name = "原始數量")]
        public int? OriQty { get; set; }
        /// <summary>新增人員</summary>
        [StringLength(10)]
        [Display(Name = "新增人員")]
        public string AddName { get; set; }
        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }
        /// <summary>最後修改人員</summary>
        [StringLength(10)]
        [Display(Name = "最後修改人員")]
        public string EditName { get; set; }
        /// <summary>最後修改時間</summary>
        [Display(Name = "最後修改時間")]
        public DateTime? EditDate { get; set; }

    }
}
