using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(FinalInspection_Detail) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/10 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/10   1.00    Admin        Create
    /// </history>
    public class FinalInspection_Detail
    {
        /// <summary>Ukey</summary>
        [Required]
        [Display(Name = "Ukey")]
        public string Ukey { get; set; }

        /// <summary>單號</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "單號")]
        public string ID { get; set; }

        /// <summary>Defect Type</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "Defect Type")]
        public string GarmentDefectTypeID { get; set; }

        /// <summary>Defect Code</summary>
        [Required]
        [StringLength(3)]
        [Display(Name = "Defect Code")]
        public string GarmentDefectCodeID { get; set; }

        /// <summary>瑕疵數量</summary>
        [Required]
        [Display(Name = "瑕疵數量")]
        public int? Qty { get; set; }

    }
}
