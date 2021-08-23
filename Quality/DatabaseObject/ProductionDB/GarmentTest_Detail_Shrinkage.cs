using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(GarmentTest_Detail_Shrinkage) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>
    public class GarmentTest_Detail_Shrinkage
    {
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public Int64 ID { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public int? No { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "")]
        public string Location { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(150)]
        [Display(Name = "")]
        public string Type { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal BeforeWash { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal SizeSpec { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal AfterWash1 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Shrinkage1 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal AfterWash2 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Shrinkage2 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal AfterWash3 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Shrinkage3 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Seq { get; set; }

    }
}
