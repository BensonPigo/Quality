using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(MockupCrocking) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/19 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/19   1.00    Admin        Create
    /// </history>
    public class MockupCrocking2
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
        /// <summary>測試日期</summary>
        [Display(Name = "測試日期")]
        public DateTime? TestDate { get; set; }

        public DateTime? ReceivedDate { get; set; }

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
        /// <summary>區分大貨階段 (B) 與開發階段 (S)</summary>
        [Display(Name = "區分大貨階段 (B) 與開發階段 (S)")]
        public string Type { get; set; }
        /// <summary>測試前拍照</summary>
        [Display(Name = "測試前拍照")]
        public Byte[] TestBeforePicture { get; set; }
        /// <summary>測試後拍照</summary>
        [Display(Name = "測試後拍照")]
        public Byte[] TestAfterPicture { get; set; }
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
    }


    public class MockupCrocking
    {
        public string ReportNo { get; set; }

        private string _POID;

        public string POID
        {
            get => _POID ?? string.Empty;
            set => _POID = value;
        }

        private string _StyleID;

        public string StyleID
        {
            get => _StyleID ?? string.Empty;
            set => _StyleID = value;
        }

        private string _SeasonID;

        public string SeasonID
        {
            get => _SeasonID ?? string.Empty;
            set => _SeasonID = value;
        }

        private string _BrandID;

        public string BrandID
        {
            get => _BrandID ?? string.Empty;
            set => _BrandID = value;
        }

        private string _Article;

        public string Article
        {
            get => _Article ?? string.Empty;
            set => _Article = value;
        }

        private string _ArtworkTypeID;

        public string ArtworkTypeID
        {
            get => _ArtworkTypeID ?? string.Empty;
            set => _ArtworkTypeID = value;
        }

        private string _Remark;

        public string Remark
        {
            get => _Remark ?? string.Empty;
            set => _Remark = value;
        }

        private string _T1Subcon;

        public string T1Subcon
        {
            get => _T1Subcon ?? string.Empty;
            set => _T1Subcon = value;
        }

        public DateTime? TestDate { get; set; }


        public DateTime? ReceivedDate { get; set; }


        public DateTime? ReleasedDate { get; set; }

        private string _Result;

        public string Result
        {
            get => _Result ?? string.Empty;
            set => _Result = value;
        }

        private string _Technician;

        public string Technician
        {
            get => _Technician ?? string.Empty;
            set => _Technician = value;
        }

        private string _MR;

        public string MR
        {
            get => _MR ?? string.Empty;
            set => _MR = value;
        }

        private string _Type;

        public string Type
        {
            get => _Type ?? string.Empty;
            set => _Type = value;
        }

        public Byte[] TestBeforePicture { get; set; }

        public Byte[] TestAfterPicture { get; set; }


        public DateTime? AddDate { get; set; }

        private string _AddName;

        public string AddName
        {
            get => _AddName ?? string.Empty;
            set => _AddName = value;
        }


        public DateTime? EditDate { get; set; }

        private string _EditName;

        public string EditName
        {
            get => _EditName ?? string.Empty;
            set => _EditName = value;
        }
    }

}
