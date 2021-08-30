using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(GarmentTest_Detail_FGPT) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>
    public class GarmentTest_Detail_FGPT
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
        public string TestName { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string TestDetail { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public int? Criteria { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string TestResult { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string TestUnit { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public int? Seq { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public int? TypeSelection_VersionID { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public int? TypeSelection_Seq { get; set; }

    }
}
