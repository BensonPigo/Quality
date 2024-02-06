using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class NewCDCode
    {
        /// <summary></summary>
        [Required]
        [StringLength(15)]
        
        public string Classifty { get; set; }
        /// <summary></summary>
        [StringLength(50)]
        
        public string TypeName { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(1)]
        
        public string ID { get; set; }
        /// <summary></summary>
        [StringLength(30)]
        
        public string Placket { get; set; }
        /// <summary></summary>
        [StringLength(100)]
        
        public string Definition { get; set; }
        /// <summary></summary>
        
        public decimal CPU { get; set; }
        /// <summary></summary>
        
        public decimal ComboPcs { get; set; }
        /// <summary></summary>
        [StringLength(200)]
        
        public string Remark { get; set; }
        /// <summary></summary>
        
        public bool Junk { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        
        public string AddName { get; set; }
        /// <summary></summary>
        
        public DateTime? AddDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        
        public string EditName { get; set; }
        /// <summary></summary>
        
        public DateTime? EditDate { get; set; }

    }
}
