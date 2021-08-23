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
        [Display(Name = "")]
        public long Ukey { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "")]
        public string ID { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "")]
        public string Article { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public long FinalInspection_OrderCartonUkey { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string Instrument { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string Fabrication { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal GarmentTop { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal GarmentMiddle { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal GarmentBottom { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal CTNInside { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal CTNOutside { get; set; }
        /// <summary></summary>
        [StringLength(1)]
        [Display(Name = "")]
        public string Result { get; set; }
        /// <summary></summary>
        [StringLength(50)]
        [Display(Name = "")]
        public string Action { get; set; }
        /// <summary></summary>
        [StringLength(-1)]
        [Display(Name = "")]
        public string Remark { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string AddName { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? AddDate { get; set; }

    }
}
