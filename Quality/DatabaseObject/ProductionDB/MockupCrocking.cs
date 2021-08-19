using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(MockupCrocking) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/19 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/19   1.00    Admin        Create
    /// </history>
    public class MockupCrocking
    {
        /// <summary>測試單號</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "測試單號")]
        public string ReportNo { get; set; }
        /// <summary>訂單號碼</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "訂單號碼")]
        public string POID { get; set; }
        /// <summary>款式</summary>
        [Required]
        [StringLength(15)]
        [Display(Name = "款式")]
        public string StyleID { get; set; }
        /// <summary>季節</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "季節")]
        public string SeasonID { get; set; }
        /// <summary>品牌</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "品牌")]
        public string BrandID { get; set; }
        /// <summary>色組</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "色組")]
        public string Article { get; set; }
        /// <summary>工段</summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "工段")]
        public string ArtworkTypeID { get; set; }
        /// <summary>備註</summary>
        [Required]
        [StringLength(300)]
        [Display(Name = "備註")]
        public string Remark { get; set; }
        /// <summary>T1 廠商</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "T1 廠商")]
        public string T1Subcon { get; set; }
        /// <summary>測試日期</summary>
        [Display(Name = "測試日期")]
        public DateTime? TestDate { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? ReceivedDate { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? ReleasedDate { get; set; }
        /// <summary>檢驗結果</summary>
        [Required]
        [StringLength(4)]
        [Display(Name = "檢驗結果")]
        public string Result { get; set; }
        /// <summary>技術人員</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "技術人員")]
        public string Technician { get; set; }
        /// <summary>業務</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "業務")]
        public string MR { get; set; }
        /// <summary>區分大貨階段 (B) 與開發階段 (S)</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "區分大貨階段 (B) 與開發階段 (S)")]
        public string Type { get; set; }
        /// <summary>測試前拍照</summary>
        [StringLength(-1)]
        [Display(Name = "測試前拍照")]
        public byte[] TestBeforePicture { get; set; }
        /// <summary>測試後拍照</summary>
        [StringLength(-1)]
        [Display(Name = "測試後拍照")]
        public byte[] TestAfterPicture { get; set; }
        /// <summary>新增日期</summary>
        [Display(Name = "新增日期")]
        public DateTime? AddDate { get; set; }
        /// <summary>新增人員</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "新增人員")]
        public string AddName { get; set; }
        /// <summary>編輯日期</summary>
        [Display(Name = "編輯日期")]
        public DateTime? EditDate { get; set; }
        /// <summary>編輯人員</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "編輯人員")]
        public string EditName { get; set; }

    }
}
