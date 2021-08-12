using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(FinalInspection_OrderCarton) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/10 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/10   1.00    Admin        Create
    /// </history>
    public class FinalInspection_OrderCarton
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

        /// <summary>訂單SP</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "訂單SP")]
        public string OrderID { get; set; }

        /// <summary>Packinglist ID</summary>
        [StringLength(13)]
        [Display(Name = "Packinglist ID")]
        public string PackinglistID { get; set; }

        /// <summary>箱號</summary>
        [StringLength(6)]
        [Display(Name = "箱號")]
        public string CTNNo { get; set; }

    }
}
