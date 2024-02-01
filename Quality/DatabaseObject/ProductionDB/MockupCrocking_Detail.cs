using DatabaseObject.Public;
using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class MockupCrocking_Detail2 : CompareBase
    {
        [Display(Name = "測試單號")]
        public string ReportNo { get; set; }

        public long Ukey { get; set; }

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
    public class MockupCrocking_Detail : CompareBase
    {
        public string ReportNo { get; set; }

        public long Ukey { get; set; }

        public string Design { get; set; }

        private string _ArtworkColor;

        public string ArtworkColor
        {
            get => _ArtworkColor ?? string.Empty;
            set => _ArtworkColor = value;
        }

        private string _FabricRefNo;

        public string FabricRefNo
        {
            get => _FabricRefNo ?? string.Empty;
            set => _FabricRefNo = value;
        }

        private string _FabricColor;

        public string FabricColor
        {
            get => _FabricColor ?? string.Empty;
            set => _FabricColor = value;
        }

        private string _DryScale;

        public string DryScale
        {
            get => _DryScale ?? string.Empty;
            set => _DryScale = value;
        }

        private string _WetScale;

        public string WetScale
        {
            get => _WetScale ?? string.Empty;
            set => _WetScale = value;
        }

        public string Result { get; set; }

        public string Remark { get; set; }

        public string EditName { get; set; }

        public DateTime? EditDate { get; set; }

    }

}
