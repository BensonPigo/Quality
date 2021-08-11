using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Brand
    {
        /// <summary>客戶代碼</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "客戶代碼")]
        public string ID { get; set; }
        /// <summary>中文全名</summary>
        [StringLength(30)]
        [Display(Name = "中文全名")]
        public string NameCH { get; set; }
        /// <summary>英文全名</summary>
        [StringLength(40)]
        [Display(Name = "英文全名")]
        public string NameEN { get; set; }
        /// <summary>國別代碼</summary>
        [StringLength(2)]
        [Display(Name = "國別代碼")]
        public string CountryID { get; set; }
        /// <summary>客戶</summary>
        [StringLength(8)]
        [Display(Name = "客戶")]
        public string BuyerID { get; set; }
        /// <summary>電話</summary>
        [StringLength(20)]
        [Display(Name = "電話")]
        public string Tel { get; set; }
        /// <summary>傳真</summary>
        [StringLength(20)]
        [Display(Name = "傳真")]
        public string Fax { get; set; }
        /// <summary>聯絡人1</summary>
        [StringLength(20)]
        [Display(Name = "聯絡人1")]
        public string Contact1 { get; set; }
        /// <summary>聯絡人2</summary>
        [StringLength(20)]
        [Display(Name = "聯絡人2")]
        public string Contact2 { get; set; }
        /// <summary>中文地址</summary>
        [StringLength(50)]
        [Display(Name = "中文地址")]
        public string AddressCH { get; set; }
        /// <summary>英文地址</summary>
        [StringLength(-1)]
        [Display(Name = "英文地址")]
        public string AddressEN { get; set; }
        /// <summary>交易幣別</summary>
        [StringLength(3)]
        [Display(Name = "交易幣別")]
        public string CurrencyID { get; set; }
        /// <summary>備註</summary>
        [StringLength(-1)]
        [Display(Name = "備註")]
        public string Remark { get; set; }
        /// <summary>訂單上的自訂欄位1</summary>
        [StringLength(12)]
        [Display(Name = "訂單上的自訂欄位1")]
        public string Customize1 { get; set; }
        /// <summary>訂單上的自訂欄位2</summary>
        [StringLength(12)]
        [Display(Name = "訂單上的自訂欄位2")]
        public string Customize2 { get; set; }
        /// <summary>訂單上的自訂欄位3</summary>
        [StringLength(12)]
        [Display(Name = "訂單上的自訂欄位3")]
        public string Customize3 { get; set; }
        /// <summary>佣金%</summary>
        [Display(Name = "佣金%")]
        public Int16 Commission { get; set; }
        /// <summary>郵遞區號</summary>
        [StringLength(6)]
        [Display(Name = "郵遞區號")]
        public string ZipCode { get; set; }
        /// <summary>E_Mail Address</summary>
        [StringLength(50)]
        [Display(Name = "E_Mail Address")]
        public string Email { get; set; }
        /// <summary>業務組別</summary>
        [StringLength(5)]
        [Display(Name = "業務組別")]
        public string MrTeam { get; set; }
        /// <summary>Brand Group (用於Color,Frabic,Inventory)</summary>
        [StringLength(8)]
        [Display(Name = "Brand Group (用於Color,Frabic,Inventory)")]
        public string BrandGroup { get; set; }
        /// <summary>Apparel 範本名稱</summary>
        [StringLength(20)]
        [Display(Name = "Apparel 範本名稱")]
        public string ApparelXlt { get; set; }
        /// <summary>Sample 耗損%(Fabric)</summary>
        [Display(Name = "Sample 耗損%(Fabric)")]
        public decimal LossSampleFabric { get; set; }
        /// <summary>Payment Term for Bulk</summary>
        [StringLength(10)]
        [Display(Name = "Payment Term for Bulk")]
        public string PayTermARIDBulk { get; set; }
        /// <summary>Payment Term for Sample</summary>
        [StringLength(10)]
        [Display(Name = "Payment Term for Sample")]
        public string PayTermARIDSample { get; set; }
        /// <summary>工廠AreaCode名稱</summary>
        [StringLength(12)]
        [Display(Name = "工廠AreaCode名稱")]
        public string BrandFactoryAreaCaption { get; set; }
        /// <summary>工廠FactoryCode名稱</summary>
        [StringLength(12)]
        [Display(Name = "工廠FactoryCode名稱")]
        public string BrandFactoryCodeCaption { get; set; }
        /// <summary>工廠VendorCode名稱</summary>
        [StringLength(12)]
        [Display(Name = "工廠VendorCode名稱")]
        public string BrandFactoryVendorCaption { get; set; }
        /// <summary>大貨出貨三碼代碼(用於Invoice)</summary>
        [StringLength(3)]
        [Display(Name = "大貨出貨三碼代碼(用於Invoice)")]
        public string ShipCode { get; set; }
        /// <summary>取消</summary>
        [Display(Name = "取消")]
        public bool Junk { get; set; }
        /// <summary>新增人員</summary>
        [StringLength(10)]
        [Display(Name = "新增人員")]
        public string AddName { get; set; }
        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }
        /// <summary>最後修改人員</summary>
        [StringLength(10)]
        [Display(Name = "最後修改人員")]
        public string EditName { get; set; }
        /// <summary>最後修改時間</summary>
        [Display(Name = "最後修改時間")]
        public DateTime? EditDate { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal LossSampleAccessory { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string ShipLeader { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? ShipLeaderEditDate { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public int? OTDExtension { get; set; }
        /// <summary></summary>
        [StringLength(1)]
        [Display(Name = "")]
        public string UseRatioRule { get; set; }
        /// <summary></summary>
        [StringLength(1)]
        [Display(Name = "")]
        public string UseRatioRule_Thick { get; set; }

    }
}
