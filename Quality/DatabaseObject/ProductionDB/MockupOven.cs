using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class MockupOven
    {
        [Display(Name = "測試單號")]
        public string ReportNo { get; set; }

        [Display(Name = "訂單號碼")]
        public string POID { get; set; }

        [Display(Name = "款式")]
        public string StyleID { get; set; }

        [Display(Name = "季節")]
        public string SeasonID { get; set; }

        [Display(Name = "品牌")]
        public string BrandID { get; set; }

        [Display(Name = "色組")]
        public string Article { get; set; }

        [Display(Name = "工段")]
        public string ArtworkTypeID { get; set; }

        [Display(Name = "備註")]
        public string Remark { get; set; }

        [Display(Name = "T1 廠商")]
        public string T1Subcon { get; set; }

        [Display(Name = "T2 廠商")]
        public string T2Supplier { get; set; }

        [Display(Name = "測試日期")]
        public DateTime? TestDate { get; set; }

        [Display(Name = "")]
        public DateTime? ReceivedDate { get; set; }

        [Display(Name = "")]
        public DateTime? ReleasedDate { get; set; }

        [Display(Name = "檢驗結果")]
        public string Result { get; set; }

        [Display(Name = "技術人員")]
        public string Technician { get; set; }

        [Display(Name = "業務")]
        public string MR { get; set; }

        [Display(Name = "新增日期")]
        public DateTime? AddDate { get; set; }

        [Display(Name = "新增人員")]
        public string AddName { get; set; }

        [Display(Name = "編輯日期")]
        public DateTime? EditDate { get; set; }

        [Display(Name = "編輯人員")]
        public string EditName { get; set; }

        [Display(Name = "TestTemperature")]
        public decimal? TestTemperature { get; set; }

        [Display(Name = "TestTime")]
        public decimal? TestTime { get; set; }

        [Display(Name = "HTPlate")]
        public int? HTPlate { get; set; }

        [Display(Name = "HTFlim")]
        public int? HTFlim { get; set; }

        [Display(Name = "HTTime")]
        public int? HTTime { get; set; }

        [Display(Name = "HTPressure")]
        public decimal? HTPressure { get; set; }

        [StringLength(5)]
        [Display(Name = "HTPellOff")]
        public string HTPellOff { get; set; }

        [Display(Name = "HT2ndPressnoreverse")]
        public int? HT2ndPressnoreverse { get; set; }

        [Display(Name = "HT2ndPressreversed")]
        public int? HT2ndPressreversed { get; set; }

        [Display(Name = "HTCoolingTime")]
        public int? HTCoolingTime { get; set; }

        [Display(Name = "測試前拍照")]
        public Byte[] TestBeforePicture { get; set; }

        [Display(Name = "測試後拍照")]
        public Byte[] TestAfterPicture { get; set; }

        [Display(Name = "區分大貨階段 (B) 與開發階段 (S)")]
        public string Type { get; set; }

    }
}
