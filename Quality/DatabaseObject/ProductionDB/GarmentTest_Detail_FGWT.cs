using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(GarmentTest_Detail_FGWT) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>
    public class GarmentTest_Detail_FGWT2
    {
        /// <summary></summary>
        
        public long ID { get; set; }
        /// <summary></summary>
        
        public int No { get; set; }
        /// <summary></summary>
        
        public string Location { get; set; }
        /// <summary></summary>
        
        public string Type { get; set; }
        /// <summary></summary>
        
        public string TestDetail { get; set; }
        /// <summary></summary>
        
        public decimal BeforeWash { get; set; }
        /// <summary></summary>
        
        public decimal SizeSpec { get; set; }
        /// <summary></summary>
        
        public decimal AfterWash { get; set; }
        /// <summary></summary>
        
        public decimal Shrinkage { get; set; }
        /// <summary></summary>
        
        public string Scale { get; set; }
        /// <summary></summary>
        
        public decimal Criteria { get; set; }
        /// <summary>Pass 標準；Criteria2 主要用在判斷一個區間</summary>
        [Display(Name = "Pass 標準；Criteria2 主要用在判斷一個區間")]
        public decimal Criteria2 { get; set; }
        /// <summary></summary>
        
        public string SystemType { get; set; }

        public string StandardRemark { get; set; }
        /// <summary></summary>
        
        public int Seq { get; set; }

    }

    public class GarmentTest_Detail_FGWT
    {
        public long ID { get; set; }

        public int No { get; set; }

        private string _Location;

        public string Location
        {
            get => _Location ?? string.Empty;
            set => _Location = value;
        }

        private string _Type;

        public string Type
        {
            get => _Type ?? string.Empty;
            set => _Type = value;
        }

        private string _TestDetail;

        public string TestDetail
        {
            get => _TestDetail ?? string.Empty;
            set => _TestDetail = value;
        }

        public decimal BeforeWash { get; set; }

        public decimal SizeSpec { get; set; }

        public decimal AfterWash { get; set; }

        public decimal Shrinkage { get; set; }

        private string _Scale;

        public string Scale
        {
            get => _Scale ?? string.Empty;
            set => _Scale = value;
        }

        public decimal Criteria { get; set; }

        public decimal Criteria2 { get; set; }

        private string _SystemType;

        public string SystemType
        {
            get => _SystemType ?? string.Empty;
            set => _SystemType = value;
        }

        public string StandardRemark { get; set; }
        
        public string BaseKey { get; set; }
        public string KeyToShrinkage { get; set; }
        public int Seq { get; set; }
    }

}
