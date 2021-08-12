using System;
using System.ComponentModel.DataAnnotations;
namespace MICS.Parameter
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
        /// <summary>��˭p��</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "��˭p��")]
        public string InspectionLevels { get; set; }
        /// <summary>���~�`�� - �P�_�d��_�l��</summary>
        [Required]
        [Display(Name = "���~�`�� - �P�_�d��_�l��")]
        public int? LotSize_Start { get; set; }
        /// <summary>���~�`�� - �P�_�d�򵲧��</summary>
        [Required]
        [Display(Name = "���~�`�� - �P�_�d�򵲧��")]
        public int? LotSize_End { get; set; }
        /// <summary>��˼�</summary>
        [Required]
        [Display(Name = "��˼�")]
        public int? SampleSize { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public string Ukey { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public string Junk { get; set; }
        /// <summary>AQL類型</summary>
        [Required]
        [Display(Name = "AQL類型")]
        public string AQLType { get; set; }
        /// <summary>可容忍檢驗失敗數量</summary>
        [Display(Name = "可容忍檢驗失敗數量")]
        public int? AcceptedQty { get; set; }

    }
}
