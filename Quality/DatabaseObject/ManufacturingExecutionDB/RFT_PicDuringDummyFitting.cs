using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class RFT_PicDuringDummyFitting
    {
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
        /// <summary>尺寸</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "尺寸")]
        public string Size { get; set; }
        /// <summary>前</summary>
        [StringLength(-1)]
        [Display(Name = "前")]
        public Byte[] Front { get; set; }
        /// <summary>側</summary>
        [StringLength(-1)]
        [Display(Name = "側")]
        public Byte[] Side { get; set; }
        /// <summary>後</summary>
        [StringLength(-1)]
        [Display(Name = "後")]
        public Byte[] Back { get; set; }

    }
}
