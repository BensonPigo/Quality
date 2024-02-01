using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class SewingLine
    {
        /// <summary>產線代號</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "產線代號")]
        public string ID { get; set; }
        /// <summary>說明</summary>
        [Required]
        [StringLength(500)]
        [Display(Name = "說明")]
        public string Description { get; set; }
        /// <summary>工廠代號</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "工廠代號")]
        public string FactoryID { get; set; }
        /// <summary>產線組別</summary>
        [StringLength(2)]
        [Display(Name = "產線組別")]
        public string SewingCell { get; set; }
        /// <summary>預設車縫人數</summary>
        [Display(Name = "預設車縫人數")]
        public int Sewer { get; set; }
        /// <summary></summary>
        
        public bool Junk { get; set; }
        /// <summary>新增人員</summary>
        [StringLength(10)]
        [Display(Name = "新增人員")]
        public string AddName { get; set; }
        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }
        /// <summary>最後修改人員</summary>
        [StringLength(10)]
        [Display(Name = "最後修改人員")]
        public string EditName { get; set; }
        /// <summary>最後修改時間</summary>
        [Display(Name = "最後修改時間")]
        public DateTime? EditDate { get; set; }
        /// <summary>�̫�����ɶ�</summary>
        [Display(Name = "�̫�����ɶ�")]
        public DateTime? LastInpsectionTime { get; set; }
        /// <summary>���m�ɶ�(����)</summary>
        [Required]
        [Display(Name = "���m�ɶ�(����)")]
        public int IdleTime { get; set; }
        /// <summary>���P����ܥu��ܦ�Group�U��Line</summary>
        [StringLength(50)]
        [Display(Name = "���P����ܥu��ܦ�Group�U��Line")]
        public string LineGroup { get; set; }

    }
}
