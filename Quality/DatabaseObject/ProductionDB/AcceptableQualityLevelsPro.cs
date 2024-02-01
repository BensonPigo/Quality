using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ProductionDB
{

    public class AcceptableQualityLevelsPro
    {
        public string BrandID { get; set; }
        public string Category { get; set; }
        /// <summary>InspectionLevels</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "InspectionLevels")]
        public string InspectionLevels { get; set; }
        /// <summary>LotSize_Start</summary>
        [Required]
        [Display(Name = "LotSize_Start")]
        public int LotSize_Start { get; set; }
        /// <summary>LotSize_End</summary>
        [Required]
        [Display(Name = "LotSize_End")]
        public int LotSize_End { get; set; }
        /// <summary>SampleSize</summary>
        [Required]
        [Display(Name = "SampleSize")]
        public int SampleSize { get; set; }
        /// <summary>Ukey</summary>
        [Required]
        [Display(Name = "ProUkey")]
        public long ProUkey { get; set; }
        /// <summary>Junk</summary>
        [Required]
        [Display(Name = "Junk")]
        public bool Junk { get; set; }
        /// <summary>AQL類型</summary>
        [Required]
        [Display(Name = "AQL類型")]
        public decimal AQLType { get; set; }
        /// <summary>可容忍檢驗失敗數量</summary>
        [Display(Name = "可容忍檢驗失敗數量")]
        public int AcceptedQty { get; set; }

    }

    public class AcceptableQualityLevelsProList
    {
        public long ProUkey { get; set; }
        public string BrandID { get; set; }
        public string Category { get; set; }
        public int LotSize_Start { get; set; }
        public int LotSize_End { get; set; }
        public int SampleSize { get; set; }
        public long AQLDefectCategoryUkey { get; set; }

        /// <summary>可容忍檢驗失敗數量</summary>
        public int AcceptedQty { get; set; }
        public int RejectQty { get; set; }
        public string DefectDescription { get; set; }

        /// <summary>
        /// 以下兩個暫時還沒用到
        /// </summary>
        public string InspectionLevels { get; set; }
        public decimal AQLType { get; set; }
    }
}
