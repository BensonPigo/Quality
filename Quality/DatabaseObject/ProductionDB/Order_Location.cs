using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Order_Location
    {
        /// <summary></summary>
        [Required]
        [StringLength(13)]
        public string OrderId { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(1)]
        public string Location { get; set; }
        /// <summary></summary>
        public decimal Rate { get; set; }
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
