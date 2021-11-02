using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Quality_Menu
    {
        /// <summary>MenuID</summary>
        [Required]
        [Display(Name = "MenuID")]
        public Int64 ID { get; set; }

        /// <summary>模組名稱</summary>
        [Required]
        [StringLength(40)]
        [Display(Name = "模組名稱")]
        public string ModuleName { get; set; }

        /// <summary>模組順序</summary>
        [Required]
        [Display(Name = "模組順序")]
        public int? ModuleSeq { get; set; }

        /// <summary>功能名稱</summary>
        [Required]
        [StringLength(60)]
        [Display(Name = "功能名稱")]
        public string FunctionName { get; set; }

        /// <summary>功能順序</summary>
        [Required]
        [Display(Name = "功能順序")]
        public int? FunctionSeq { get; set; }

        /// <summary>移除</summary>
        [Required]
        [Display(Name = "移除")]
        public bool Junk { get; set; }

        [Display(Name = "Url")]
        public string Url { get; set; }
    }
}
