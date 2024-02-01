using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class TPEPass1
    {
        /// <summary>編號</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "編號")]
        public string ID { get; set; }
        /// <summary>名字</summary>
        [StringLength(50)]
        [Display(Name = "名字")]
        public string Name { get; set; }
        /// <summary>分機</summary>
        [StringLength(6)]
        [Display(Name = "分機")]
        public string ExtNo { get; set; }
        /// <summary></summary>
        [StringLength(80)]
        
        public string EMail { get; set; }

    }
}
