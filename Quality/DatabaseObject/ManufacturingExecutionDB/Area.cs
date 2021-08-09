using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Area
    {
        /// <summary></summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "")]
        public string Code { get; set; }
        /// <summary></summary>
        [StringLength(100)]
        [Display(Name = "")]
        public string Description { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool Junk { get; set; }
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
        [StringLength(7)]
        [Display(Name = "")]
        public string Type { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool T { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool B { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool I { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool O { get; set; }
        /// <summary></summary>
        [StringLength(100)]
        [Display(Name = "")]
        public string LocalCode { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public Byte Seq { get; set; }

    }
}
