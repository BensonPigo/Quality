using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Quality_Menu_Detail
    {
        /// <summary>Quality_Menu</summary>
        [Required]
        [Display(Name = "根據條件顯示不同的功能名稱")]
        public Int64 ID { get; set; }

        /// <summary>根據條件顯示不同的功能名稱</summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "根據條件顯示不同的功能名稱")]
        public string Type { get; set; }

        /// <summary>功能名稱</summary>
        [Required]
        [StringLength(60)]
        [Display(Name = "功能名稱")]
        public string FunctionName { get; set; }

    }
}
