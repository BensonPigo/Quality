using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*Garment Test(GarmentTest) 詳細敘述如下*/
    /// <summary>
    /// Garment Test
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>
    public class GarmentTest
    {
        /// <summary>ID</summary>
        [Display(Name = "ID")]
        public Int64 ID { get; set; }

        /// <summary>首次訂單號碼</summary>
        [Display(Name = "首次訂單號碼")]
        public string FirstOrderID { get; set; }

        /// <summary>檢驗訂單號碼</summary>
        [Display(Name = "檢驗訂單號碼")]
        public string OrderID { get; set; }

        /// <summary>款示</summary>
        [Display(Name = "款示")]
        public string StyleID { get; set; }

        /// <summary>季節</summary>
        [Display(Name = "季節")]
        public string SeasonID { get; set; }

        /// <summary>廠牌</summary>
        [Display(Name = "廠牌")]
        public string BrandID { get; set; }

        /// <summary>色組</summary>
        [Display(Name = "色組")]
        public string Article { get; set; }

        /// <summary>工廠代號</summary>
        [Display(Name = "工廠代號")]
        public string MDivisionid { get; set; }

        /// <summary>測試截止日</summary>
        [Display(Name = "測試截止日")]
        public DateTime? DeadLine { get; set; }

        /// <summary>Sewing 上線日</summary>
        [Display(Name = "Sewing 上線日")]
        public DateTime? SewingInline { get; set; }

        /// <summary>Sewing 下線日</summary>
        [Display(Name = "Sewing 下線日")]
        public DateTime? SewingOffline { get; set; }

        /// <summary>最後測試日期</summary>
        [Display(Name = "最後測試日期")]
        public DateTime? Date { get; set; }

        /// <summary>結果</summary>
        [Display(Name = "結果")]
        public string Result { get; set; }

        /// <summary>備註</summary>
        [Display(Name = "備註")]
        public string Remark { get; set; }

        /// <summary>新增人員</summary>
        [Display(Name = "新增人員")]
        public string AddName { get; set; }

        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }

        /// <summary>最後編輯人員</summary>
        [Display(Name = "最後編輯人員")]
        public string EditName { get; set; }

        /// <summary>最後編輯時間</summary>
        [Display(Name = "最後編輯時間")]
        public DateTime? EditDate { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string OldUkey { get; set; }

        /// <summary>Adidas-PHX-AP0450 SeamBrakage 最後一次檢驗的結果</summary>
        [Display(Name = "Adidas-PHX-AP0450 SeamBrakage 最後一次檢驗的結果")]
        public string SeamBreakageResult { get; set; }

        /// <summary>Adidas-PHX-AP0450 SeamBrakage 最後一次檢驗的日期</summary>
        [Display(Name = "Adidas-PHX-AP0450 SeamBrakage 最後一次檢驗的日期")]
        public DateTime? SeamBreakageLastTestDate { get; set; }

        /// <summary>Adidas-PHX-AP0451 最後一次檢驗的結果</summary>
        [Display(Name = "Adidas-PHX-AP0451 最後一次檢驗的結果")]
        public string OdourResult { get; set; }

        /// <summary>Adidas-PHX-AP0701 / PHX-AP0710 最後一次檢驗的結果</summary>
        [Display(Name = "Adidas-PHX-AP0701 / PHX-AP0710 最後一次檢驗的結果")]
        public string WashResult { get; set; }

    }
}
