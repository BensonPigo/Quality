using DatabaseObject.Public;
using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class MockupCrocking_Detail : CompareBase
    {
        [Display(Name = "測試單號")]
        public string ReportNo { get; set; }

        public Int64 Ukey { get; set; }

        public string Design { get; set; }

        [Display(Name = "工段顏色")]
        public string ArtworkColor { get; set; }

        [Display(Name = "主料料號")]
        public string FabricRefNo { get; set; }

        [Display(Name = "主料顏色")]
        public string FabricColor { get; set; }

        [Display(Name = "乾摩擦色牢度評級")]
        public string DryScale { get; set; }

        [Display(Name = "濕摩擦色牢度評級")]
        public string WetScale { get; set; }

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
