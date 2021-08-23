using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(Garment_Detail_Spirality) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>
    public class Garment_Detail_Spirality
    {
        /// <summary>Garment_Detail.ID</summary>
        [Required]
        [Display(Name = "Garment_Detail.ID")]
        public long ID { get; set; }

        /// <summary>Garment_Detail.No</summary>
        [Required]
        [Display(Name = "Garment_Detail.No")]
        public int? No { get; set; }

        /// <summary>���� (Top, Bottom)</summary>
        [Required]
        [StringLength(1)]
        public string Location { get; set; }

        [Required]
        public decimal MethodA_AAPrime { get; set; }

        [Required]
        public decimal MethodA_APrimeB { get; set; }

        [Required]
        public decimal MethodB_AAPrime { get; set; }

        [Required]
        public decimal MethodB_AB { get; set; }

        [Required]
        public decimal CM { get; set; }

        [Required]
        public decimal MethodA { get; set; }

        [Required]
        public decimal MethodB { get; set; }

    }
}
