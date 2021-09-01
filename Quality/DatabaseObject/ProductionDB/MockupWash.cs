using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(MockupWash) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/31 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/31   1.00    Admin        Create
    /// </history>
    public class MockupWash
    {
        /// <summary>測試單號</summary>
        [Display(Name = "測試單號")]
        public string ReportNo { get; set; }
        /// <summary>訂單號碼</summary>
        [Display(Name = "訂單號碼")]
        public string POID { get; set; }
        /// <summary>款式</summary>
        [Display(Name = "款式")]
        public string StyleID { get; set; }
        /// <summary>季節</summary>
        [Display(Name = "季節")]
        public string SeasonID { get; set; }
        /// <summary>品牌</summary>
        [Display(Name = "品牌")]
        public string BrandID { get; set; }
        /// <summary>色組</summary>
        [Display(Name = "色組")]
        public string Article { get; set; }
        /// <summary>工段</summary>
        [Display(Name = "工段")]
        public string ArtworkTypeID { get; set; }
        /// <summary>備註</summary>
        [Display(Name = "備註")]
        public string Remark { get; set; }
        /// <summary>T1 廠商</summary>
        [Display(Name = "T1 廠商")]
        public string T1Subcon { get; set; }
        /// <summary>T2 廠商</summary>
        [Display(Name = "T2 廠商")]
        public string T2Supplier { get; set; }
        /// <summary>測試日期</summary>
        [Display(Name = "測試日期")]
        public DateTime? TestDate { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? ReceivedDate { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? ReleasedDate { get; set; }
        /// <summary>檢驗結果</summary>
        [Display(Name = "檢驗結果")]
        public string Result { get; set; }
        /// <summary>技術人員</summary>
        [Display(Name = "技術人員")]
        public string Technician { get; set; }
        /// <summary>業務</summary>
        [Display(Name = "業務")]
        public string MR { get; set; }
        /// <summary>新增日期</summary>
        [Display(Name = "新增日期")]
        public DateTime? AddDate { get; set; }
        /// <summary>新增人員</summary>
        [Display(Name = "新增人員")]
        public string AddName { get; set; }
        /// <summary>編輯日期</summary>
        [Display(Name = "編輯日期")]
        public DateTime? EditDate { get; set; }
        /// <summary>編輯人員</summary>
        [Display(Name = "編輯人員")]
        public string EditName { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool OtherMethod { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public string MethodID { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public string TestingMethod { get; set; }
        /// <summary>HTPlate</summary>
        [Display(Name = "HTPlate")]
        public int? HTPlate { get; set; }
        /// <summary>HTFlim</summary>
        [Display(Name = "HTFlim")]
        public int? HTFlim { get; set; }
        /// <summary>HTTime</summary>
        [Display(Name = "HTTime")]
        public int? HTTime { get; set; }
        /// <summary>HTPressure</summary>
        [Display(Name = "HTPressure")]
        public decimal HTPressure { get; set; }
        /// <summary>HTPellOff</summary>
        [Display(Name = "HTPellOff")]
        public string HTPellOff { get; set; }
        /// <summary>HT2ndPressnoreverse</summary>
        [Display(Name = "HT2ndPressnoreverse")]
        public int? HT2ndPressnoreverse { get; set; }
        /// <summary>HT2ndPressreversed</summary>
        [Display(Name = "HT2ndPressreversed")]
        public int? HT2ndPressreversed { get; set; }
        /// <summary>HTCoolingTime</summary>
        [Display(Name = "HTCoolingTime")]
        public string HTCoolingTime { get; set; }
        /// <summary>測試前拍照</summary>
        [Display(Name = "測試前拍照")]
        public Byte[] TestBeforePicture { get; set; }
        /// <summary>測試後拍照</summary>
        [Display(Name = "測試後拍照")]
        public Byte[] TestAfterPicture { get; set; }
        /// <summary>區分大貨階段 (B) 與開發階段 (S)</summary>
        [Display(Name = "區分大貨階段 (B) 與開發階段 (S)")]
        public string Type { get; set; }

    }
}
