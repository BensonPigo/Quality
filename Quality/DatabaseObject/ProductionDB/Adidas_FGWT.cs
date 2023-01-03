using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
   
    public class Adidas_FGWT
    {
        [Required]
        public int? Seq { get; set; }

        [Required]
        [StringLength(15)]
        public string TestName { get; set; }

        [Required]
        [StringLength(1)]
        public string Location { get; set; }

        [Required]
        [StringLength(150)]
        public string SystemType { get; set; }

        [Required]
        [StringLength(150)]
        public string ReportType { get; set; }

        [Required]
        [StringLength(20)]
        public string MtlTypeID { get; set; }

        [Required]
        [StringLength(20)]
        public string Washing { get; set; }
        
        [Required]
        [StringLength(30)]
        public string FabricComposition { get; set; }
        public string StandardRemark { get; set; }

        [Required]
        [StringLength(30)]
        public string TestDetail { get; set; }

        [StringLength(5)]
        public string Scale { get; set; }
        
        public decimal Criteria { get; set; }

        public string Criteria2 { get; set; }

    }
}
