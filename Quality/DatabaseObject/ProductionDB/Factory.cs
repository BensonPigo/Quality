using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Factory
    {
        /// <summary>工廠代號</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "工廠代號")]
        public string ID { get; set; }

        /// <summary>組織代號</summary>
        [StringLength(8)]
        [Display(Name = "組織代號")]
        public string MDivisionID { get; set; }

        /// <summary>取消</summary>
        [Display(Name = "取消")]
        public bool Junk { get; set; }

        /// <summary>工廠簡稱</summary>
        [StringLength(10)]
        [Display(Name = "工廠簡稱")]
        public string Abb { get; set; }

        /// <summary>工廠名稱</summary>
        [StringLength(40)]
        [Display(Name = "工廠名稱")]
        public string NameCH { get; set; }

        /// <summary>英文名稱</summary>
        [StringLength(40)]
        [Display(Name = "英文名稱")]
        public string NameEN { get; set; }

        /// <summary>國別</summary>
        [StringLength(2)]
        [Display(Name = "國別")]
        public string CountryID { get; set; }

        /// <summary>電話</summary>
        [StringLength(30)]
        [Display(Name = "電話")]
        public string Tel { get; set; }

        /// <summary>傳真</summary>
        [StringLength(30)]
        [Display(Name = "傳真")]
        public string Fax { get; set; }

        /// <summary>中文地址</summary>
        [StringLength(50)]
        [Display(Name = "中文地址")]
        public string AddressCH { get; set; }

        /// <summary>英文地址</summary>
        [Display(Name = "英文地址")]
        public string AddressEN { get; set; }

        /// <summary>幣別</summary>
        [StringLength(3)]
        [Display(Name = "幣別")]
        public string CurrencyID { get; set; }

        /// <summary>月總產值</summary>
        [Display(Name = "月總產值")]
        public int? CPU { get; set; }

        /// <summary>郵遞區號</summary>
        [StringLength(6)]
        [Display(Name = "郵遞區號")]
        public string ZipCode { get; set; }

        /// <summary>PORT OF SEA(First sale)</summary>
        [StringLength(20)]
        [Display(Name = "PORT OF SEA(First sale)")]
        public string PortSea { get; set; }

        /// <summary>Port of AIR (First AIR)</summary>
        [StringLength(20)]
        [Display(Name = "Port of AIR (First AIR)")]
        public string PortAir { get; set; }

        /// <summary>Production Kits group</summary>
        [StringLength(8)]
        [Display(Name = "Production Kits group")]
        public string KitId { get; set; }

        /// <summary>國際快遞group</summary>
        [StringLength(8)]
        [Display(Name = "國際快遞group")]
        public string ExpressGroup { get; set; }

        /// <summary>IE operation 編碼代號</summary>
        [StringLength(1)]
        [Display(Name = "IE operation 編碼代號")]
        public string IECode { get; set; }

        /// <summary>Region Code用於Nego 產生</summary>
        [StringLength(3)]
        [Display(Name = "Region Code用於Nego 產生")]
        public string NegoRegion { get; set; }

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

        /// <summary>工廠kpi統計群組</summary>
        [StringLength(8)]
        [Display(Name = "工廠kpi統計群組")]
        public string KPICode { get; set; }

        /// <summary>Factory Group</summary>
        [StringLength(3)]
        [Display(Name = "Factory Group")]
        public string FTYGroup { get; set; }

        /// <summary>Key Word</summary>
        [StringLength(3)]
        [Display(Name = "Key Word")]
        public string KeyWord { get; set; }

        /// <summary>稅籍編號</summary>
        [StringLength(15)]
        [Display(Name = "稅籍編號")]
        public string TINNo { get; set; }

        /// <summary>增值稅會計科目</summary>
        [StringLength(8)]
        [Display(Name = "增值稅會計科目")]
        public string VATAccNo { get; set; }

        /// <summary>預扣稅會計科目</summary>
        [StringLength(8)]
        [Display(Name = "預扣稅會計科目")]
        public string WithholdingRateAccNo { get; set; }

        /// <summary>扣款銀行會計科目</summary>
        [StringLength(8)]
        [Display(Name = "扣款銀行會計科目")]
        public string CreditBankAccNo { get; set; }

        /// <summary>廠長</summary>
        [StringLength(10)]
        [Display(Name = "廠長")]
        public string Manager { get; set; }

        /// <summary>使用APS系統</summary>
        [Display(Name = "使用APS系統")]
        public bool UseAPS { get; set; }

        /// <summary>使用Subcon Bundle Tracking System</summary>
        [Display(Name = "使用Subcon Bundle Tracking System")]
        public bool UseSBTS { get; set; }

        /// <summary>Garment booking amend時是否需要檢查出口報關申請</summary>
        [Display(Name = "Garment booking amend時是否需要檢查出口報關申請")]
        public bool CheckDeclare { get; set; }

        /// <summary>Factory CMT包含Local Purchase</summary>
        [Display(Name = "Factory CMT包含Local Purchase")]
        public bool LocalCMT { get; set; }

        /// <summary>Type :Bulk ,Sample ,MMS</summary>
        [StringLength(1)]
        [Display(Name = "Type :Bulk ,Sample ,MMS")]
        public string Type { get; set; }

        /// <summary>工廠地區別</summary>
        [StringLength(6)]
        [Display(Name = "工廠地區別")]
        public string Zone { get; set; }

        /// <summary>工廠排序碼</summary>
        [StringLength(3)]
        [Display(Name = "工廠排序碼")]
        public string FactorySort { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public bool IsSampleRoom { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public bool IsSCI { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public bool IsProduceFty { get; set; }

        /// <summary></summary>
        [StringLength(8)]
        [Display(Name = "")]
        public string TestDocFactoryGroup { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public bool IsOriginalFty { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? LastDownloadAPSDate { get; set; }

        /// <summary></summary>
        [StringLength(8)]
        [Display(Name = "")]
        public bool FtyZone { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string Foundry { get; set; }

        /// <summary></summary>
        [StringLength(8)]
        public string ProduceM { get; set; }

        /// <summary></summary>
        [StringLength(8)]
        public string LoadingFactoryGroup { get; set; }

    }
}
