using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Style_Artwork
    {
        [Display(Name = "款式的唯一值")]
        public long StyleUkey { get; set; }

        [Display(Name = "作工種類")]
        public string ArtworkTypeID { get; set; }

        [Display(Name = "顏色組")]
        public string Article { get; set; }

        [Display(Name = "裁片代號")]
        public string PatternCode { get; set; }

        [Display(Name = "裁片名稱")]
        public string PatternDesc { get; set; }

        [Display(Name = "印繡花模號或板號")]
        public string ArtworkID { get; set; }

        [Display(Name = "印繡花名稱")]
        public string ArtworkName { get; set; }

        [Display(Name = "數量")]
        public int Qty { get; set; }

        [Display(Name = "報價")]
        public decimal Price { get; set; }

        [Display(Name = "成本")]
        public decimal Cost { get; set; }

        [Display(Name = "備註")]
        public string Remark { get; set; }

        [Display(Name = "Ukey")]
        public long Ukey { get; set; }

        [Display(Name = "新增人員")]
        public string AddName { get; set; }

        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }

        [Display(Name = "最後修改人員")]
        public string EditName { get; set; }

        [Display(Name = "最後修改時間")]
        public DateTime? EditDate { get; set; }

        
        public int TMS { get; set; }

        
        public long TradeUkey { get; set; }

        
        public string SMNoticeID { get; set; }

        
        public string PatternVersion { get; set; }

        
        public int ActStitch { get; set; }

        
        public decimal PPU { get; set; }

        
        public string InkType { get; set; }

        
        public string Colors { get; set; }

        
        public decimal Length { get; set; }

        
        public decimal Width { get; set; }

        
        public bool AntiMigration { get; set; }

        
        public string PrintType { get; set; }
    }
}
