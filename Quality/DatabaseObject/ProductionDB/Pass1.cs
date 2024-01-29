using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Pass1
    {
        /// <summary></summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "")]
        public string ID { get; set; }

        /// <summary></summary>
        [StringLength(30)]
        [Display(Name = "")]
        public string Name { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string Password { get; set; }

        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string Position { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public Int64 FKPass0 { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public bool IsAdmin { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public bool IsMIS { get; set; }

        /// <summary></summary>
        [StringLength(2)]
        [Display(Name = "")]
        public string OrderGroup { get; set; }

        /// <summary></summary>
        [StringLength(50)]
        [Display(Name = "")]
        public string EMail { get; set; }

        /// <summary></summary>
        [StringLength(6)]
        [Display(Name = "")]
        public string ExtNo { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? OnBoard { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? Resign { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string Supervisor { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string Manager { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string Deputy { get; set; }

        /// <summary></summary>
        [StringLength(100)]
        [Display(Name = "")]
        public string Factory { get; set; }

        /// <summary></summary>
        [StringLength(6)]
        [Display(Name = "")]
        public string CodePage { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string AddName { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? AddDate { get; set; }

        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string EditName { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? EditDate { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? LastLoginTime { get; set; }

        /// <summary></summary>
        [StringLength(60)]
        [Display(Name = "")]
        public string ESignature { get; set; }

        /// <summary></summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "")]
        public string Remark { get; set; }
        public string ADAccount { get; set; }

    }
}
