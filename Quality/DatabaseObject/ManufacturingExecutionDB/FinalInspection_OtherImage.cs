using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(FinalInspection_OtherImage) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/16 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/16   1.00    Admin        Create
    /// </history>
    public class FinalInspection_OtherImage
    {
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public long Ukey { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "")]
        public string ID { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(-1)]
        [Display(Name = "")]
        public byte[] Image { get; set; }

    }
}
