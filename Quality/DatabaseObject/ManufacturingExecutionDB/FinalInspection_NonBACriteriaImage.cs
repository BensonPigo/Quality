using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(FinalInspection_NonBACriteriaImage) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/16 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/16   1.00    Admin        Create
    /// </history>
    public class FinalInspection_NonBACriteriaImage
    {
        /// <summary></summary>
        [Required]
        
        public long Ukey { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(13)]
        
        public string ID { get; set; }
        /// <summary></summary>
        [Required]
        
        public long FinalInspection_NonBACriteriaUkey { get; set; }
        /// <summary></summary>
        [Required]
        
        public byte[] Image { get; set; }

    }
}
