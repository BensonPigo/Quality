using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(FinalInspection_Measurement) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/10 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/10   1.00    Admin        Create
    /// </history>
    public class FinalInspection_Measurement
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

        /// <summary>色組</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "色組")]
        public string Article { get; set; }

        /// <summary>尺寸</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "尺寸")]
        public string SizeCode { get; set; }

        /// <summary>部位(T/B/IO)</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "部位(T/B/IO)")]
        public string Location { get; set; }

        /// <summary>編號</summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "編號")]
        public string Code { get; set; }

        /// <summary>尺寸規格</summary>
        [Required]
        [StringLength(15)]
        [Display(Name = "尺寸規格")]
        public string SizeSpec { get; set; }

        /// <summary>MeasurementUkey</summary>
        [Required]
        [Display(Name = "MeasurementUkey")]
        public long MeasurementUkey { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        
        public string AddName { get; set; }

        /// <summary></summary>
        
        public DateTime? AddDate { get; set; }

    }
}
