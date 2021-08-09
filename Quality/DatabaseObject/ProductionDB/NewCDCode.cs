using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class NewCDCode
    {
        /// <summary></summary>
        [Required]
        [StringLength(15)]
        [Display(Name = "")]
        public string Classifty { get; set; }
        /// <summary></summary>
        [StringLength(50)]
        [Display(Name = "")]
        public string TypeName { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "")]
        public string ID { get; set; }
        /// <summary></summary>
        [StringLength(30)]
        [Display(Name = "")]
        public string Placket { get; set; }
        /// <summary></summary>
        [StringLength(100)]
        [Display(Name = "")]
        public string Definition { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal CPU { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal ComboPcs { get; set; }
        /// <summary></summary>
        [StringLength(200)]
        [Display(Name = "")]
        public string Remark { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool Junk { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string AddName { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? AddDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string EditName { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? EditDate { get; set; }

    }
}
