using System.ComponentModel.DataAnnotations;

namespace DatabaseObject.ManufacturingExecutionDB
{
    public class ReworkList
    {
        /// <summary></summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "")]
        public string ReworkNo { get; set; }

        [Required]
        [StringLength(13)]
        [Display(Name = "")]
        public string SPNo { get; set; }

        [Required]
        [StringLength(30)]
        [Display(Name = "")]
        public string PONo { get; set; }

        [Required]
        [StringLength(15)]
        [Display(Name = "")]
        public string StyleNo { get; set; }

        [Required]
        [StringLength(15)]
        [Display(Name = "")]
        public string Size { get; set; }

        [Required]
        [StringLength(8)]
        [Display(Name = "")]
        public string Article { get; set; }

        [Required]
        [StringLength(120)]
        [Display(Name = "")]
        public string Defect { get; set; }

        [Required]
        [StringLength(12)]
        [Display(Name = "")]
        public string Reworked { get; set; }

        [Required]
        [StringLength(20)]
        [Display(Name = "")]
        public string AddDefect { get; set; }

        [Required]
        [StringLength(10)]
        [Display(Name = "")]
        public string Status { get; set; }

        [Required]
        [StringLength(8)]
        [Display(Name = "")]
        public string FactoryID { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "")]
        public string Line { get; set; }

        [Required]
        [StringLength(5)]
        [Display(Name = "")]
        public string ReworkCardType { get; set; }

        [Required]
        [StringLength(2)]
        [Display(Name = "")]
        public string ReworkCardNo { get; set; }
    }
}
