using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class GarmentTest_Detail_Shrinkage2
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

        public decimal BeforeWash { get; set; }
        /// <summary></summary>

        public decimal SizeSpec { get; set; }
        /// <summary></summary>

        public decimal AfterWash1 { get; set; }
        /// <summary></summary>

        public decimal Shrinkage1 { get; set; }
        /// <summary></summary>

        public decimal AfterWash2 { get; set; }
        /// <summary></summary>

        public decimal Shrinkage2 { get; set; }
        /// <summary></summary>

        public decimal AfterWash3 { get; set; }
        /// <summary></summary>

        public decimal Shrinkage3 { get; set; }
        /// <summary></summary>

        public decimal Seq { get; set; }

    }
    public class GarmentTest_Detail_Shrinkage
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

        public decimal BeforeWash { get; set; }

        public decimal SizeSpec { get; set; }

        public decimal AfterWash1 { get; set; }

        public decimal Shrinkage1 { get; set; }

        public decimal AfterWash2 { get; set; }


        public decimal Shrinkage2 { get; set; }


        public decimal AfterWash3 { get; set; }


        public decimal Shrinkage3 { get; set; }
        public decimal Seq { get; set; }
        public string BaseKey { get; set; }
        public string KeyToFgwt { get; set; }
}

}
