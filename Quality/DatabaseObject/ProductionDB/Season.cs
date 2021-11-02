using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*各Brand的季別基本檔(Season) 詳細敘述如下*/
    /// <summary>
    /// 各Brand的季別基本檔
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/19 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/19   1.00    Admin        Create
    /// </history>
    public class Season
    {
        /// <summary>Brand 的季別</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "Brand 的季別")]
        public string ID { get; set; }
        /// <summary>Brand</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "Brand")]
        public string BrandID { get; set; }
        /// <summary>Material Alarm Cost%</summary>
        [Display(Name = "Material Alarm Cost%")]
        public string CostRatio { get; set; }
        /// <summary>Sci Season</summary>
        [StringLength(10)]
        [Display(Name = "Sci Season")]
        public string SeasonSCIID { get; set; }
        /// <summary>Brand Season 的起始年月</summary>
        [StringLength(7)]
        [Display(Name = "Brand Season 的起始年月")]
        public string Month { get; set; }
        /// <summary>取消</summary>
        [Display(Name = "取消")]
        public string Junk { get; set; }
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
