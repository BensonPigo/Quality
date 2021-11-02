using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*Grey Scale 基本檔(Scale) 詳細敘述如下*/
    /// <summary>
    /// Grey Scale 基本檔
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/30 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/30   1.00    Admin        Create
    /// </history>
    public class Scale
    {
        /// <summary>Scale 代號</summary>
        [Required]
        [StringLength(5)]
        [Display(Name = "Scale 代號")]
        public string ID { get; set; }
        /// <summary>取消</summary>
        [Display(Name = "取消")]
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

    }
}
