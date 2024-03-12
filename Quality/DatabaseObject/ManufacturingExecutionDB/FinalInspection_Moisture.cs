using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(FinalInspection_Moisture) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/16 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/16   1.00    Admin        Create
    /// </history>
    public class FinalInspection_Moisture
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
        [StringLength(8)]
        
        public string Article { get; set; }
        /// <summary></summary>
        
        public long FinalInspection_OrderCartonUkey { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        
        public string Instrument { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        
        public string Fabrication { get; set; }
        /// <summary></summary>
        
        public decimal GarmentTop { get; set; }
        /// <summary></summary>
        
        public decimal GarmentMiddle { get; set; }
        /// <summary></summary>
        
        public decimal GarmentBottom { get; set; }
        /// <summary></summary>
        
        public decimal CTNInside { get; set; }
        /// <summary></summary>
        
        public decimal CTNOutside { get; set; }
        /// <summary></summary>
        [StringLength(1)]
        
        public string Result { get; set; }
        /// <summary></summary>
        [StringLength(50)]
        
        public string Action { get; set; }
        /// <summary></summary>
        
        public string Remark { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        
        public string AddName { get; set; }
        /// <summary></summary>
        
        public DateTime? AddDate { get; set; }

    }
}
