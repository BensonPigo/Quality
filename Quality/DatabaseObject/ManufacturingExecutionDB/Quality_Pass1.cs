using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Quality_Pass1
    {
        /// <summary>使用者 ID</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "使用者 ID")]
        public string ID { get; set; }
        /// <summary>帳號來源 (PMS, MES)</summary>
        [Required]
        [StringLength(3)]
        [Display(Name = "帳號來源 (PMS, MES)")]
        public string ImportFrom { get; set; }
        /// <summary>姓名</summary>
        [Required]
        [StringLength(30)]
        [Display(Name = "姓名")]
        public string Name { get; set; }
        /// <summary>密碼</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "密碼")]
        public string Password { get; set; }
        /// <summary>職位</summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "職位")]
        public string Position { get; set; }
        /// <summary>Email</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "Email")]
        public string Email { get; set; }

    }
}
