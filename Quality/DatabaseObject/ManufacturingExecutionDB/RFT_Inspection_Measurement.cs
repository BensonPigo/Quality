using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class RFT_Inspection_Measurement
    {
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public Int64 Ukey { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public Int64 MeasurementUkey { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public Int64 StyleUkey { get; set; }
        /// <summary>第幾次檢驗</summary>
        [Required]
        [Display(Name = "第幾次檢驗")]
        public int? No { get; set; }
        /// <summary>代號</summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "代號")]
        public string Code { get; set; }
        /// <summary>尺碼</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "尺碼")]
        public string SizeCode { get; set; }
        /// <summary>測量後的結果</summary>
        [Required]
        [StringLength(15)]
        [Display(Name = "測量後的結果")]
        public string SizeSpec { get; set; }
        /// <summary>訂單號碼</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "訂單號碼")]
        public string OrderID { get; set; }
        /// <summary>色組</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "色組")]
        public string Article { get; set; }
        /// <summary>部位</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "部位")]
        public string Location { get; set; }
        /// <summary>產線</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "產線")]
        public string Line { get; set; }
        /// <summary>工廠</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "工廠")]
        public string FactoryID { get; set; }

    }
}
