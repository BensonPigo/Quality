using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*Style Article(Style_Article) 詳細敘述如下*/
    /// <summary>
    /// Style Article
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/19 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/19   1.00    Admin        Create
    /// </history>
    public class Style_Article
    {
        /// <summary>款式的唯一值</summary>
        [Required]
        [Display(Name = "款式的唯一值")]
        public string StyleUkey { get; set; }
        /// <summary>序號</summary>
        [Display(Name = "序號")]
        public string Seq { get; set; }
        /// <summary>顏色組</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "顏色組")]
        public string Article { get; set; }
        /// <summary>棉紙</summary>
        [Display(Name = "棉紙")]
        public string TissuePaper { get; set; }
        /// <summary>顏色組名稱</summary>
        [StringLength(100)]
        [Display(Name = "顏色組名稱")]
        public string ArticleName { get; set; }
        /// <summary>成衣成份</summary>
        [StringLength(-1)]
        [Display(Name = "成衣成份")]
        public string Contents { get; set; }

    }
}
