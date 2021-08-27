using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class GarmentTest_Detail_Shrinkage
    {
        /// <summary></summary>
        [Display(Name = "")]
        public Int64? ID { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public int? No { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public string Location { get; set; }
        /// <summary></summary>
        [StringLength(150)]
        [Display(Name = "")]
        public string Type { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal BeforeWash { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal SizeSpec { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal AfterWash1 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Shrinkage1 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal AfterWash2 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Shrinkage2 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal AfterWash3 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Shrinkage3 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Seq { get; set; }

    }
}
