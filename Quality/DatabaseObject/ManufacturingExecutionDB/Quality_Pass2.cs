using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Quality_Pass2
    {
        /// <summary>職務</summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "職務")]
        public string PositionID { get; set; }
        /// <summary>MenuID</summary>
        [Required]
        [Display(Name = "MenuID")]
        public Int64 MenuID { get; set; }
        /// <summary>可否使用</summary>
        [Required]
        [Display(Name = "可否使用")]
        public bool Used { get; set; }

    }
}
