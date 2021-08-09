using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class RFT_Inspection_Detail
    {
        /// <summary>檢驗紀錄 ID (CFT_Inspeciton.ID)</summary>
        [Required]
        [Display(Name = "檢驗紀錄 ID (CFT_Inspeciton.ID)")]
        public Int64 ID { get; set; }
        /// <summary>檢驗瑕疵紀錄</summary>
        [Required]
        [Display(Name = "檢驗瑕疵紀錄")]
        public Int64 Ukey { get; set; }
        /// <summary>瑕疵描述</summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "瑕疵描述")]
        public string DefectCode { get; set; }
        /// <summary>瑕疵位置</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "瑕疵位置")]
        public string AreaCode { get; set; }
        /// <summary>Junk</summary>
        [Required]
        [Display(Name = "Junk")]
        public bool Junk { get; set; }
        /// <summary>Beautiful Audit Criteria ID</summary>
        [Required]
        [StringLength(3)]
        [Display(Name = "Beautiful Audit Criteria ID")]
        public string PMS_RFTBACriteriaID { get; set; }
        /// <summary>責任歸屬 ID</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "責任歸屬 ID")]
        public string PMS_RFTRespID { get; set; }
        /// <summary>瑕疵代號</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "瑕疵代號")]
        public string GarmentDefectTypeID { get; set; }
        /// <summary>瑕疵代號</summary>
        [Required]
        [StringLength(3)]
        [Display(Name = "瑕疵代號")]
        public string GarmentDefectCodeID { get; set; }
        /// <summary>瑕疵圖片</summary>
        [StringLength(-1)]
        [Display(Name = "瑕疵圖片")]
        public Byte[] DefectPicture { get; set; }
        /// <summary>新增時間</summary>
        [Required]
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }

    }
}
