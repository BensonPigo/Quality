using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class MailTo
    {
        /// <summary></summary>
        [Required]
        [StringLength(3)]
        
        public string ID { get; set; }

        /// <summary></summary>
        [StringLength(60)]
        
        public string Description { get; set; }

        /// <summary></summary>
        [StringLength(5000)]
        
        public string ToAddress { get; set; }

        /// <summary></summary>
        [StringLength(5000)]
        
        public string CcAddress { get; set; }

        /// <summary></summary>
        [StringLength(100)]
        
        public string Subject { get; set; }

        /// <summary></summary>
        [StringLength(5000)]
        
        public string Content { get; set; }

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
