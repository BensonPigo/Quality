using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(EndlineMoisture) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/16 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/16   1.00    Admin        Create
    /// </history>
    public class EndlineMoisture
    {
        /// <summary>�˴�����</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "�˴�����")]
        public string Instrument { get; set; }
        /// <summary>���������զ�</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "���������զ�")]
        public string Fabrication { get; set; }
        /// <summary>�зǭ�</summary>
        [Display(Name = "�зǭ�")]
        public decimal Standard { get; set; }
        /// <summary>�зǭȳ��</summary>
        [StringLength(1)]
        [Display(Name = "�зǭȳ��")]
        public string Unit { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool Junk { get; set; }
        /// <summary></summary>
        [StringLength(500)]
        [Display(Name = "")]
        public string Description { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? AddDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string AddName { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? EditDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string Editname { get; set; }

    }
}
