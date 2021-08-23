using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(FinalInspection_NonBACriteria) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/10 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/10   1.00    Admin        Create
    /// </history>
    public class FinalInspection_NonBACriteria
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

        /// <summary>評比項目</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "評比項目")]
        public string BACriteria { get; set; }

        /// <summary>不好看數量</summary>
        [Required]
        [Display(Name = "不好看數量")]
        public int? Qty { get; set; }

    }
}
