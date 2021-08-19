using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(MockupCrocking_Detail) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/19 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/19   1.00    Admin        Create
    /// </history>
    public class MockupCrocking_Detail
    {
        /// <summary>測試單號</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "測試單號")]
        public string ReportNo { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public long Ukey { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "")]
        public string Design { get; set; }
        /// <summary>工段顏色</summary>
        [Required]
        [StringLength(35)]
        [Display(Name = "工段顏色")]
        public string ArtworkColor { get; set; }
        /// <summary>主料料號</summary>
        [Required]
        [StringLength(30)]
        [Display(Name = "主料料號")]
        public string FabricRefNo { get; set; }
        /// <summary>主料顏色</summary>
        [Required]
        [StringLength(35)]
        [Display(Name = "主料顏色")]
        public string FabricColor { get; set; }
        /// <summary>乾摩擦色牢度評級</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "乾摩擦色牢度評級")]
        public string DryScale { get; set; }
        /// <summary>濕摩擦色牢度評級</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "濕摩擦色牢度評級")]
        public string WetScale { get; set; }
        /// <summary>測試結果</summary>
        [Required]
        [StringLength(4)]
        [Display(Name = "測試結果")]
        public string Result { get; set; }
        /// <summary>備註</summary>
        [Required]
        [StringLength(300)]
        [Display(Name = "備註")]
        public string Remark { get; set; }
        /// <summary>編輯人員</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "編輯人員")]
        public string EditName { get; set; }
        /// <summary>編輯日期</summary>
        [Display(Name = "編輯日期")]
        public DateTime? EditDate { get; set; }

    }
}
