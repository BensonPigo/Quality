using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Quality_Position
    {
        /// <summary>職務</summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "職務")]
        public string ID { get; set; }
        /// <summary>描述</summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "描述")]
        public string Description { get; set; }
        /// <summary>是否為系統管理員</summary>
        [Required]
        [Display(Name = "是否為系統管理員")]
        public bool IsAdmin { get; set; }
        /// <summary>移除</summary>
        [Required]
        [Display(Name = "移除")]
        public bool Junk { get; set; }

    }
}
