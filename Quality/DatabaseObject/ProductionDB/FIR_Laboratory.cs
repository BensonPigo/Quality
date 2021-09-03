using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class FIR_Laboratory
    {
        /// <summary>id</summary>
        [Display(Name = "id")]
        public Int64? ID { get; set; }

        /// <summary>採購單號</summary>
        [Display(Name = "採購單號")]
        public string POID { get; set; }

        /// <summary>大小項</summary>
        [Display(Name = "大小項")]
        public string SEQ1 { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string SEQ2 { get; set; }

        /// <summary>檢驗截止日</summary>
        [Display(Name = "檢驗截止日")]
        public DateTime? InspDeadline { get; set; }

        /// <summary>色脫落結果</summary>
        [Display(Name = "色脫落結果")]
        public string Crocking { get; set; }

        /// <summary>熱壓縮律結果</summary>
        [Display(Name = "熱壓縮律結果")]
        public string Heat { get; set; }

        /// <summary>水洗縮律結果</summary>
        [Display(Name = "水洗縮律結果")]
        public string Wash { get; set; }

        /// <summary>色脫落測試 日期</summary>
        [Display(Name = "色脫落測試 日期")]
        public DateTime? CrockingDate { get; set; }

        /// <summary>熱縮律測試日期</summary>
        [Display(Name = "熱縮律測試日期")]
        public DateTime? HeatDate { get; set; }

        /// <summary>水洗縮律測試日期</summary>
        [Display(Name = "水洗縮律測試日期")]
        public DateTime? WashDate { get; set; }

        /// <summary>色脫落備註</summary>
        [Display(Name = "色脫落備註")]
        public string CrockingRemark { get; set; }

        /// <summary>熱縮律測試備註</summary>
        [Display(Name = "熱縮律測試備註")]
        public string HeatRemark { get; set; }

        /// <summary>水洗縮律測是備註</summary>
        [Display(Name = "水洗縮律測是備註")]
        public string WashRemark { get; set; }

        /// <summary>樣品收到日</summary>
        [Display(Name = "樣品收到日")]
        public DateTime? ReceiveSampleDate { get; set; }

        /// <summary>總結果</summary>
        [Display(Name = "總結果")]
        public string Result { get; set; }

        /// <summary>免測試色脫落</summary>
        [Display(Name = "免測試色脫落")]
        public bool nonCrocking { get; set; }

        /// <summary>免測試熱縮</summary>
        [Display(Name = "免測試熱縮")]
        public bool nonHeat { get; set; }

        /// <summary>免測試水洗縮</summary>
        [Display(Name = "免測試水洗縮")]
        public bool nonWash { get; set; }

        /// <summary>色脫落確認</summary>
        [Display(Name = "色脫落確認")]
        public bool CrockingEncode { get; set; }

        /// <summary>熱縮確認</summary>
        [Display(Name = "熱縮確認")]
        public bool HeatEncode { get; set; }

        /// <summary>水洗確認</summary>
        [Display(Name = "水洗確認")]
        public bool WashEncode { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string SkewnessOptionID { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string CrockingInspector { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string HeatInspector { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string WashInspector { get; set; }

        /// <summary>摩擦測試前的照片</summary>
        [Display(Name = "摩擦測試前的照片")]
        public Byte[] CrockingTestBeforePicture { get; set; }

        /// <summary>摩擦測試後的照片</summary>
        [Display(Name = "摩擦測試後的照片")]
        public Byte[] CrockingTestAfterPicture { get; set; }

        /// <summary>烘箱測試前的照片</summary>
        [Display(Name = "烘箱測試前的照片")]
        public Byte[] HeatTestBeforePicture { get; set; }

        /// <summary>烘箱測試後的照片</summary>
        [Display(Name = "烘箱測試後的照片")]
        public Byte[] HeatTestAfterPicture { get; set; }

        /// <summary>水洗測試前的照片</summary>
        [Display(Name = "水洗測試前的照片")]
        public Byte[] WashTestBeforePicture { get; set; }

        /// <summary>水洗測試後的照片</summary>
        [Display(Name = "水洗測試後的照片")]
        public Byte[] WashTestAfterPicture { get; set; }

    }

    public class FIR_Laboratory_Utility
    {
        public static readonly string UpdateResultSql = @"
update  FIR_Laboratory set  Result = case when  (Crocking = 'Pass' or nonCrocking = 1) and 
                                                (Heat = 'Pass' or nonHeat = 1) and 
                                                (Wash = 'Pass' or nonWash = 1) then 'Pass'
                                          when  (Crocking = '' and nonCrocking = 0) or
                                                (Heat = '' and nonHeat = 0) or
                                                (Wash = '' and nonWash = 0) then ''
                                          else  'Fail' end
where   ID = @ID
";
    }
}
