using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Area
    {
        /// <summary></summary>
        [Required]
        [StringLength(50)]
        
        public string Code { get; set; }
        /// <summary></summary>
        [StringLength(100)]
        
        public string Description { get; set; }
        /// <summary></summary>
        
        public bool Junk { get; set; }
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
        [StringLength(7)]
        
        public string Type { get; set; }
        /// <summary></summary>
        [Required]
        
        public bool T { get; set; }
        /// <summary></summary>
        [Required]
        
        public bool B { get; set; }
        /// <summary></summary>
        [Required]
        
        public bool I { get; set; }
        /// <summary></summary>
        [Required]
        
        public bool O { get; set; }
        /// <summary></summary>
        [StringLength(100)]
        
        public string LocalCode { get; set; }
        /// <summary></summary>
        [Required]
        
        public Byte Seq { get; set; }

    }
}
