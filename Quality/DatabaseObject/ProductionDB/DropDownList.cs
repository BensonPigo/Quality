using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class DropDownList
    {
        /// <summary>Type</summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "Type")]
        public string Type { get; set; }

        /// <summary>ID</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "ID")]
        public string ID { get; set; }

        /// <summary>Name</summary>
        [StringLength(50)]
        [Display(Name = "Name")]
        public string Name { get; set; }

        /// <summary>實際存的長度</summary>
        [Required]
        [Display(Name = "實際存的長度")]
        public decimal RealLength { get; set; }

        /// <summary>描述</summary>
        [StringLength(150)]
        [Display(Name = "描述")]
        public string Description { get; set; }

        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public int? Seq { get; set; }

    }
}
