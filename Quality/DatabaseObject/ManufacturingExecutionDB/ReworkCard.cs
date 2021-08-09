using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class ReworkCard
    {
        /// <summary></summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "")]
        public string No { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(5)]
        [Display(Name = "")]
        public string Type { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "")]
        public string FactoryID { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "")]
        public string Line { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? AddDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string AddName { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? EditDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string EditName { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(6)]
        [Display(Name = "")]
        public string Status { get; set; }

    }
}
