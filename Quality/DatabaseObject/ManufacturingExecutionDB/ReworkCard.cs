using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class ReworkCard
    {
        /// <summary></summary>
        [Required]
        [StringLength(2)]
        
        public string No { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(5)]
        
        public string Type { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(8)]
        
        public string FactoryID { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(2)]
        
        public string Line { get; set; }
        /// <summary></summary>
        
        public DateTime? AddDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        
        public string AddName { get; set; }
        /// <summary></summary>
        
        public DateTime? EditDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        
        public string EditName { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(6)]
        
        public string Status { get; set; }

    }
}
