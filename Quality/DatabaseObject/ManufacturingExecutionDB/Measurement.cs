using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Measurement
    {
        /// <summary></summary>
        [Required]
        
        public long StyleUkey { get; set; }
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
        
        public string Description { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        
        public string Code { get; set; }
        /// <summary></summary>
        [StringLength(8)]
        
        public string SizeCode { get; set; }
        /// <summary></summary>
        [StringLength(15)]
        
        public string SizeSpec { get; set; }
        /// <summary></summary>
        [Required]
        
        public long Ukey { get; set; }
        /// <summary></summary>
        
        public DateTime? AddDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        
        public string AddName { get; set; }
        /// <summary></summary>
        
        public bool Junk { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(1)]
        
        public string SizeGroup { get; set; }
        /// <summary>MeasurementTranslate Ukey</summary>
        [Display(Name = "MeasurementTranslate Ukey")]
        public long MeasurementTranslateUkey { get; set; }

        public bool IsPatternMeas { get; set; }
    }
}
