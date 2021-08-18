using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Quality_Pass1
    {
        /// <summary>使用者 ID</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "使用者 ID")]
        public string ID { get; set; }

        /// <summary>職位</summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "職位")]
        public string Position { get; set; }

        /// <summary>BulkFGT_Brand</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "BulkFGT_Brand")]
        public string BulkFGT_Brand { get; set; }

    }
}
