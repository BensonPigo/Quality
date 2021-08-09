using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Measurement
    {
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public Int64 StyleUkey { get; set; }
        /// <summary>Tol(-)</summary>
        [StringLength(15)]
        [Display(Name = "Tol(-)")]
        public string Tol1 { get; set; }
        /// <summary>Tol(+)</summary>
        [StringLength(15)]
        [Display(Name = "Tol(+)")]
        public string Tol2 { get; set; }
        /// <summary></summary>
        [StringLength(300)]
        [Display(Name = "")]
        public string Description { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string Code { get; set; }
        /// <summary></summary>
        [StringLength(8)]
        [Display(Name = "")]
        public string SizeCode { get; set; }
        /// <summary></summary>
        [StringLength(15)]
        [Display(Name = "")]
        public string SizeSpec { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public Int64 Ukey { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? AddDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string AddName { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool Junk { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "")]
        public string SizeGroup { get; set; }
        /// <summary>MeasurementTranslate Ukey</summary>
        [Display(Name = "MeasurementTranslate Ukey")]
        public Int64 MeasurementTranslateUkey { get; set; }

    }
}
