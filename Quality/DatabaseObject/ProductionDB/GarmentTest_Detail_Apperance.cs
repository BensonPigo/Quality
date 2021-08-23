using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(GarmentTest_Detail_Apperance) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>
    public class GarmentTest_Detail_Apperance
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
        [StringLength(200)]
        [Display(Name = "")]
        public string Type { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "")]
        public string Wash1 { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "")]
        public string Wash2 { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "")]
        public string Wash3 { get; set; }
        /// <summary></summary>
        [StringLength(100)]
        [Display(Name = "")]
        public string Comment { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public int? Seq { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "")]
        public string Wash4 { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "")]
        public string Wash5 { get; set; }

    }
}
