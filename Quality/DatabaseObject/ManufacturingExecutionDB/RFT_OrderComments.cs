using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class RFT_OrderComments
    {
        /// <summary>訂單號碼</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "訂單號碼")]
        public string OrderID { get; set; }
        /// <summary>Comments 類別</summary>
        [Required]
        [StringLength(3)]
        [Display(Name = "Comments 類別")]
        public string PMS_RFTCommentsID { get; set; }
        /// <summary>CFT 註解</summary>
        [Required]
        [StringLength(1000)]
        [Display(Name = "CFT 註解")]
        public string Comnments { get; set; }

    }
}
