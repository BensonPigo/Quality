using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(GarmentTest_Detail_FGWT) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>
    public class GarmentTest_Detail_FGWT
    {
        /// <summary></summary>
        [Display(Name = "")]
        public Int64? ID { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public int? No { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public string Location { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public string Type { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public string TestDetail { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal BeforeWash { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal SizeSpec { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal AfterWash { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Shrinkage { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public string Scale { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal Criteria { get; set; }
        /// <summary>Pass 標準；Criteria2 主要用在判斷一個區間</summary>
        [Display(Name = "Pass 標準；Criteria2 主要用在判斷一個區間")]
        public decimal Criteria2 { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public string SystemType { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public int? Seq { get; set; }

    }
}
