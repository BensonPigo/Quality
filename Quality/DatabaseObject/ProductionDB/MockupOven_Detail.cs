using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(MockupOven_Detail) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/09/02 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/09/02   1.00    Admin        Create
    /// </history>
    public class MockupOven_Detail
    {
        [Display(Name = "測試單號")]
        public string ReportNo { get; set; }

        [Display(Name = "")]
        public long Ukey { get; set; }

        [Display(Name = "")]
        public string TypeofPrint { get; set; }

        [Display(Name = "")]
        public string Design { get; set; }

        [Display(Name = "工段顏色")]
        public string ArtworkColor { get; set; }

        [Display(Name = "主料料號")]
        public string FabricRefNo { get; set; }

        [Display(Name = "主料顏色")]
        public string FabricColor { get; set; }

        [Display(Name = "測試結果")]
        public string Result { get; set; }

        [Display(Name = "備註")]
        public string Remark { get; set; }

        [Display(Name = "編輯人員")]
        public string EditName { get; set; }

        [Display(Name = "編輯日期")]
        public DateTime? EditDate { get; set; }

    }
}
