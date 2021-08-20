using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*���`����榸��˽�q�����з� (AQL)(AcceptableQualityLevels) 詳細敘述如下*/
    /// <summary>
    /// ���`����榸��˽�q�����з� (AQL)
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/10 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/10   1.00    Admin        Create
    /// </history>
    public class AcceptableQualityLevels
    {
        /// <summary>InspectionLevels</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "InspectionLevels")]
        public string InspectionLevels { get; set; }
        /// <summary>LotSize_Start</summary>
        [Required]
        [Display(Name = "LotSize_Start")]
        public int? LotSize_Start { get; set; }
        /// <summary>LotSize_End</summary>
        [Required]
        [Display(Name = "LotSize_End")]
        public int? LotSize_End { get; set; }
        /// <summary>SampleSize</summary>
        [Required]
        [Display(Name = "SampleSize")]
        public int? SampleSize { get; set; }
        /// <summary>Ukey</summary>
        [Required]
        [Display(Name = "Ukey")]
        public long Ukey { get; set; }
        /// <summary>Junk</summary>
        [Required]
        [Display(Name = "Junk")]
        public bool Junk { get; set; }
        /// <summary>AQL類型</summary>
        [Required]
        [Display(Name = "AQL類型")]
        public decimal AQLType { get; set; }
        /// <summary>可容忍檢驗失敗數量</summary>
        [Display(Name = "可容忍檢驗失敗數量")]
        public int? AcceptedQty { get; set; }

    }
}
