using System;
using System.ComponentModel.DataAnnotations;


namespace DatabaseObject.Public
{
    public class Window_Brand
    {
        /// <summary>客戶代碼</summary>
        [Display(Name = "客戶代碼")]
        public string ID { get; set; }
    }

    public class Window_Season
    {
        public string ID { get; set; }
        public string BrandID { get; set; }
    }

    public class Window_Style
    {
        public string ID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
    }

    public class Window_Article
    {
        public string Article { get; set; }
        public Int64 StyleUkey { get; set; }
        public string OrderID { get; set; }
    }

    public class Window_Size
    {
        public string SizeCode { get; set; }
        public string Article { get; set; }
        public Int64 StyleUkey { get; set; }
        public string OrderID { get; set; }
    }

    public class Window_Technician
    {
        public string CallFunction { get; set; }
        public string Region { get; set; }

        public string ID { get; set; }
        public string Name { get; set; }
        public string ExtNo { get; set; }
        public string Factory { get; set; }
    }
}
