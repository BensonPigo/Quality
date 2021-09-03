using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Style_BOA
    {
        [Display(Name = "款式的唯一值")]
        public long StyleUkey { get; set; }

        [Display(Name = "唯一值")]
        public long Ukey { get; set; }

        [Display(Name = "Referce No.")]
        public string Refno { get; set; }

        [Display(Name = "飛雁料號")]
        public string SCIRefno { get; set; }

        [Display(Name = "採購大項編號")]
        public string SEQ1 { get; set; }

        [Display(Name = "單件用量")]
        public decimal ConsPC { get; set; }

        [Display(Name = "部位別")]
        public string PatternPanel { get; set; }

        [Display(Name = "量法的項目")]
        public string SizeItem { get; set; }

        [Display(Name = "使用數量由樣品室提供")]
        public bool ProvidedPatternRoom { get; set; }

        [Display(Name = "備註")]
        public string Remark { get; set; }

        [Display(Name = "顏色說明")]
        public string ColorDetail { get; set; }

        [Display(Name = "客人物料展開規則")]
        public int IsCustCD { get; set; }

        [Display(Name = "計算採購時是否判斷左右插")]
        public bool BomTypeZipper { get; set; }

        [Display(Name = "依尺寸展開")]
        public bool BomTypeSize { get; set; }

        [Display(Name = "依顏色展開")]
        public bool BomTypeColor { get; set; }

        [Display(Name = "依款式展開")]
        public bool BomTypeStyle { get; set; }

        [Display(Name = "依顏色組展開")]
        public bool BomTypeArticle { get; set; }

        [Display(Name = "依客人訂單單號展開")]
        public bool BomTypePo { get; set; }

        [Display(Name = "依客戶資料展開")]
        public bool BomTypeCustCD { get; set; }

        [Display(Name = "依工廠展開")]
        public bool BomTypeFactory { get; set; }

        [Display(Name = "依月份展開")]
        public bool BomTypeBuyMonth { get; set; }

        [Display(Name = "依工廠的國家別展開")]
        public bool BomTypeCountry { get; set; }

        [Display(Name = "大貨指定的廠商")]
        public string SuppIDBulk { get; set; }

        [Display(Name = "銷樣指定的廠商")]
        public string SuppIDSample { get; set; }

        [Display(Name = "新增人員")]
        public string AddName { get; set; }

        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }

        [Display(Name = "最後修改人員")]
        public string EditName { get; set; }

        [Display(Name = "最後修改時間")]
        public DateTime? EditDate { get; set; }

        [Display(Name = "")]
        public string FabricPanelCode { get; set; }

    }
}
