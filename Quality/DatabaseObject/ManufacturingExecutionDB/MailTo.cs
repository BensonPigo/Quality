using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class MailTo
    {
        /// <summary></summary>
        [Required]
        [StringLength(3)]
        [Display(Name = "")]
        public string ID { get; set; }

        /// <summary></summary>
        [StringLength(60)]
        [Display(Name = "")]
        public string Description { get; set; }

        /// <summary></summary>
        [StringLength(5000)]
        [Display(Name = "")]
        public string ToAddress { get; set; }

        /// <summary></summary>
        [StringLength(5000)]
        [Display(Name = "")]
        public string CcAddress { get; set; }

        /// <summary></summary>
        [StringLength(100)]
        [Display(Name = "")]
        public string Subject { get; set; }

        /// <summary></summary>
        [StringLength(5000)]
        [Display(Name = "")]
        public string Content { get; set; }

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
