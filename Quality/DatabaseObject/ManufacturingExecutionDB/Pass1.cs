using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Pass1
    {
        /// <summary></summary>
        [Required]
        [StringLength(10)]
        
        public string ID { get; set; }

        /// <summary></summary>
        [StringLength(30)]
        
        public string Name { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        
        public string Password { get; set; }

        /// <summary></summary>
        [StringLength(20)]
        
        public string Position { get; set; }

        /// <summary></summary>
        
        public long FKPass0 { get; set; }

        /// <summary></summary>
        
        public bool IsAdmin { get; set; }

        /// <summary></summary>
        
        public bool IsMIS { get; set; }

        /// <summary></summary>
        [StringLength(2)]
        
        public string OrderGroup { get; set; }

        /// <summary></summary>
        [StringLength(50)]
        
        public string EMail { get; set; }

        /// <summary></summary>
        [StringLength(6)]
        
        public string ExtNo { get; set; }

        /// <summary></summary>
        
        public DateTime? OnBoard { get; set; }

        /// <summary></summary>
        
        public DateTime? Resign { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        
        public string Supervisor { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        
        public string Manager { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        
        public string Deputy { get; set; }

        /// <summary></summary>
        [StringLength(150)]
        
        public string Factory { get; set; }

        /// <summary></summary>
        [StringLength(6)]
        
        public string CodePage { get; set; }

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

        /// <summary></summary>
        
        public DateTime? LastLoginTime { get; set; }

        /// <summary></summary>
        [Required]
        [StringLength(100)]
        
        public string Remark { get; set; }

    }
}
