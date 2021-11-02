using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class FIR_Laboratory_Heat
    {
        /// <summary>ID</summary>
        [Display(Name = "ID")]
        public Int64 ID { get; set; }

        /// <summary>捲號</summary>
        [Display(Name = "捲號")]
        public string Roll { get; set; }

        /// <summary>缸號</summary>
        [Display(Name = "缸號")]
        public string Dyelot { get; set; }

        /// <summary>檢驗日期</summary>
        [Display(Name = "檢驗日期")]
        public DateTime? Inspdate { get; set; }

        /// <summary>檢驗人員</summary>
        [Display(Name = "檢驗人員")]
        public string Inspector { get; set; }

        /// <summary>檢驗結果</summary>
        [Display(Name = "檢驗結果")]
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

        /// <summary>水平縮律</summary>
        [Display(Name = "水平縮律")]
        public decimal HorizontalRate { get; set; }

        /// <summary>水平原長</summary>
        [Display(Name = "水平原長")]
        public decimal HorizontalOriginal { get; set; }

        /// <summary>水平1</summary>
        [Display(Name = "水平1")]
        public decimal HorizontalTest1 { get; set; }

        /// <summary>水平2</summary>
        [Display(Name = "水平2")]
        public decimal HorizontalTest2 { get; set; }

        /// <summary>水平3</summary>
        [Display(Name = "水平3")]
        public decimal HorizontalTest3 { get; set; }

        /// <summary>垂直縮律</summary>
        [Display(Name = "垂直縮律")]
        public decimal VerticalRate { get; set; }

        /// <summary>垂直原長</summary>
        [Display(Name = "垂直原長")]
        public decimal VerticalOriginal { get; set; }

        /// <summary>垂直1</summary>
        [Display(Name = "垂直1")]
        public decimal VerticalTest1 { get; set; }

        /// <summary>垂直2</summary>
        [Display(Name = "垂直2")]
        public decimal VerticalTest2 { get; set; }

        /// <summary>垂直3</summary>
        [Display(Name = "垂直3")]
        public decimal VerticalTest3 { get; set; }

    }
}
