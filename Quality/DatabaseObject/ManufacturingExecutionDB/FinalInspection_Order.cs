using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(FinalInspection_Order) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/10 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/10   1.00    Admin        Create
    /// </history>
    public class FinalInspection_Order
    {
        /// <summary>Ukey</summary>
        [Required]
        [Display(Name = "Ukey")]
        public long Ukey { get; set; }

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

        /// <summary>本次可檢驗的數量</summary>
        [Display(Name = "本次可檢驗的數量")]
        public int AvailableQty { get; set; }

    }
}
