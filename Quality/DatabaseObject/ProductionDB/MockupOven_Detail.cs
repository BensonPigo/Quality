using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class MockupOven_Detail2 : CompareBase
    {
        [Display(Name = "測試單號")]
        public string ReportNo { get; set; }

        public long Ukey { get; set; }

        public string TypeofPrint { get; set; }

        public string Design { get; set; }

        [Display(Name = "工段顏色")]
        public string ArtworkColor { get; set; }

        [Display(Name = "熱轉印物料")]
        public string AccessoryRefno { get; set; }

        [Display(Name = "主料料號")]
        public string FabricRefNo { get; set; }

        [Display(Name = "主料顏色")]
        public string FabricColor { get; set; }

        [Display(Name = "測試結果")]
        public string Result { get; set; }

        [Display(Name = "色差灰階")]
        public string ChangeScale { get; set; }

        [Display(Name = "色差檢驗結果")]
        public string ResultChange { get; set; }

        [Display(Name = "染色灰階")]
        public string StainingScale { get; set; }

        [Display(Name = "染色檢驗結果")]
        public string ResultStain { get; set; }

        [Display(Name = "備註")]
        public string Remark { get; set; }

        [Display(Name = "編輯人員")]
        public string EditName { get; set; }

        [Display(Name = "編輯日期")]
        public DateTime? EditDate { get; set; }

    }

    public class MockupOven_Detail : CompareBase
    {
        public string ReportNo { get; set; }

        public long Ukey { get; set; }

        public string TypeofPrint { get; set; }

        public string Design { get; set; }

        private string _ArtworkColor;

        public string ArtworkColor
        {
            get => _ArtworkColor ?? string.Empty;
            set => _ArtworkColor = value;
        }

        private string _AccessoryRefno;

        public string AccessoryRefno
        {
            get => _AccessoryRefno ?? string.Empty;
            set => _AccessoryRefno = value;
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

        public string Result { get; set; }

        private string _ChangeScale;

        public string ChangeScale
        {
            get => _ChangeScale ?? string.Empty;
            set => _ChangeScale = value;
        }

        private string _ResultChange;

        public string ResultChange
        {
            get => _ResultChange ?? string.Empty;
            set => _ResultChange = value;
        }

        private string _StainingScale;

        public string StainingScale
        {
            get => _StainingScale ?? string.Empty;
            set => _StainingScale = value;
        }

        private string _ResultStain;

        public string ResultStain
        {
            get => _ResultStain ?? string.Empty;
            set => _ResultStain = value;
        }

        public string Remark { get; set; }

        public string EditName { get; set; }

        public DateTime? EditDate { get; set; }

    }

}
