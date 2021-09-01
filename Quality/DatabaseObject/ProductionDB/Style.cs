using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Style
    {
        /// <summary>款式號碼</summary>
        [Required]
        [StringLength(15)]
        [Display(Name = "款式號碼")]
        public string ID { get; set; }
        /// <summary>唯一值</summary>
        [Required]
        [Display(Name = "唯一值")]
        public Int64 Ukey { get; set; }
        /// <summary>品牌</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "品牌")]
        public string BrandID { get; set; }
        /// <summary>系列代碼</summary>
        [Required]
        [StringLength(12)]
        [Display(Name = "系列代碼")]
        public string ProgramID { get; set; }
        /// <summary>季別</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "季別")]
        public string SeasonID { get; set; }
        /// <summary>模型</summary>
        [StringLength(25)]
        [Display(Name = "模型")]
        public string Model { get; set; }
        /// <summary>說明</summary>
        [StringLength(100)]
        [Display(Name = "說明")]
        public string Description { get; set; }
        /// <summary>款式名稱</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "款式名稱")]
        public string StyleName { get; set; }
        /// <summary>成衣套裝組合</summary>
        [StringLength(4)]
        [Display(Name = "成衣套裝組合")]
        public string ComboType { get; set; }
        /// <summary>產能代號</summary>
        [Required]
        [StringLength(6)]
        [Display(Name = "產能代號")]
        public string CdCodeID { get; set; }
        /// <summary>成衣種類</summary>
        [Required]
        [StringLength(5)]
        [Display(Name = "成衣種類")]
        public string ApparelType { get; set; }
        /// <summary>主料種類</summary>
        [Required]
        [StringLength(5)]
        [Display(Name = "主料種類")]
        public string FabricType { get; set; }
        /// <summary>成衣成份</summary>
        [Required]
        [Display(Name = "成衣成份")]
        public string Contents { get; set; }
        /// <summary>成衣LEADTIME</summary>
        [Required]
        [Display(Name = "成衣LEADTIME")]
        public Int16 GMTLT { get; set; }
        /// <summary>產能</summary>
        [Display(Name = "產能")]
        public decimal CPU { get; set; }
        /// <summary>生產工廠</summary>
        [StringLength(200)]
        [Display(Name = "生產工廠")]
        public string Factories { get; set; }
        /// <summary>Planning 放置工廠備註用</summary>
        [Display(Name = "Planning 放置工廠備註用")]
        public string FTYRemark { get; set; }
        /// <summary>銷樣階段的訂單主管</summary>
        [StringLength(10)]
        [Display(Name = "銷樣階段的訂單主管")]
        public string SampleSMR { get; set; }
        /// <summary>Sample階段的訂單Handle</summary>
        [StringLength(10)]
        [Display(Name = "Sample階段的訂單Handle")]
        public string SampleMRHandle { get; set; }
        /// <summary>大貨階段訂單Handle主管</summary>
        [StringLength(10)]
        [Display(Name = "大貨階段訂單Handle主管")]
        public string BulkSMR { get; set; }
        /// <summary>大貨階段的訂單Handle</summary>
        [StringLength(10)]
        [Display(Name = "大貨階段的訂單Handle")]
        public string BulkMRHandle { get; set; }
        /// <summary>作廢</summary>
        [Display(Name = "作廢")]
        public bool Junk { get; set; }
        /// <summary>水洗測式</summary>
        [Display(Name = "水洗測式")]
        public bool RainwearTestPassed { get; set; }
        /// <summary>尺寸</summary>
        [StringLength(2)]
        [Display(Name = "尺寸")]
        public string SizePage { get; set; }
        /// <summary>尺碼範圍</summary>
        [Display(Name = "尺碼範圍")]
        public string SizeRange { get; set; }
        /// <summary>裝箱件數</summary>
        [Display(Name = "裝箱件數")]
        public Int16 CTNQty { get; set; }
        /// <summary>採購標準成本</summary>
        [Display(Name = "採購標準成本")]
        public decimal StdCost { get; set; }
        /// <summary>加工方式</summary>
        [StringLength(60)]
        [Display(Name = "加工方式")]
        public string Processes { get; set; }
        /// <summary>加工成本建立方式</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "加工成本建立方式")]
        public string ArtworkCost { get; set; }
        /// <summary>圖片1</summary>
        [StringLength(60)]
        [Display(Name = "圖片1")]
        public string Picture1 { get; set; }
        /// <summary>圖片2</summary>
        [StringLength(60)]
        [Display(Name = "圖片2")]
        public string Picture2 { get; set; }
        /// <summary>標籤 & 保養說明標籤</summary>
        [Display(Name = "標籤 & 保養說明標籤")]
        public string Label { get; set; }
        /// <summary>包裝方式</summary>
        [Display(Name = "包裝方式")]
        public string Packing { get; set; }
        /// <summary>IE申請單號</summary>
        [StringLength(10)]
        [Display(Name = "IE申請單號")]
        public string IETMSID { get; set; }
        /// <summary>Time Study 的版本</summary>
        [StringLength(3)]
        [Display(Name = "Time Study 的版本")]
        public string IETMSVersion { get; set; }
        /// <summary>IE 更新者</summary>
        [StringLength(10)]
        [Display(Name = "IE 更新者")]
        public string IEImportName { get; set; }
        /// <summary>IE 更新日</summary>
        [Display(Name = "IE 更新日")]
        public DateTime? IEImportDate { get; set; }
        /// <summary>款式核准日</summary>
        [Display(Name = "款式核准日")]
        public DateTime? ApvDate { get; set; }
        /// <summary>款式核准人</summary>
        [StringLength(10)]
        [Display(Name = "款式核准人")]
        public string ApvName { get; set; }
        /// <summary>洗水標籤/洗水嘜</summary>
        [StringLength(8)]
        [Display(Name = "洗水標籤/洗水嘜")]
        public string CareCode { get; set; }
        /// <summary>特別註記</summary>
        [StringLength(5)]
        [Display(Name = "特別註記")]
        public string SpecialMark { get; set; }
        /// <summary>襯</summary>
        [StringLength(20)]
        [Display(Name = "襯")]
        public string Lining { get; set; }
        /// <summary>款式單位</summary>
        [StringLength(8)]
        [Display(Name = "款式單位")]
        public string StyleUnit { get; set; }
        /// <summary>Expection Form</summary>
        [Display(Name = "Expection Form")]
        public bool ExpectionForm { get; set; }
        /// <summary>Exception Form Remark</summary>
        [Display(Name = "Exception Form Remark")]
        public string ExpectionFormRemark { get; set; }
        /// <summary>工廠當地業務</summary>
        [StringLength(10)]
        [Display(Name = "工廠當地業務")]
        public string LocalMR { get; set; }
        /// <summary>工廠自行建立之款式</summary>
        [Display(Name = "工廠自行建立之款式")]
        public bool LocalStyle { get; set; }
        /// <summary>Factory PP Meeting</summary>
        [Display(Name = "Factory PP Meeting")]
        public DateTime? PPMeeting { get; set; }
        /// <summary>不需要PP Meeting</summary>
        [Display(Name = "不需要PP Meeting")]
        public bool NoNeedPPMeeting { get; set; }
        /// <summary>Sample approval</summary>
        [Display(Name = "Sample approval")]
        public DateTime? SampleApv { get; set; }
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
        [StringLength(8)]
        [Display(Name = "")]
        public string SizeUnit { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string ModularParent { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal CPUAdjusted { get; set; }
        /// <summary>目前階段</summary>
        [StringLength(10)]
        [Display(Name = "目前階段")]
        public string Phase { get; set; }
        /// <summary>性別</summary>
        [StringLength(10)]
        [Display(Name = "性別")]
        public string Gender { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string ThreadEditname { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? ThreadEditdate { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool ThickFabric { get; set; }
        /// <summary></summary>
        [StringLength(5)]
        [Display(Name = "")]
        public string DyeingID { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string TPEEditName { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? TPEEditDate { get; set; }
        /// <summary></summary>
        public int? Pressing1 { get; set; }
        /// <summary></summary>
        public int? Pressing2 { get; set; }
        /// <summary></summary>
        public int? Folding1 { get; set; }
        /// <summary></summary>
        public int? Folding2 { get; set; }
        /// <summary>Approce/Reject</summary>
        [StringLength(1)]
        [Display(Name = "Approce/Reject")]
        public string ExpectionFormStatus { get; set; }
        /// <summary></summary>
        public DateTime? ExpectionFormDate { get; set; }
        /// <summary></summary>
        public bool ThickFabricBulk { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool HangerPack { get; set; }
        /// <summary></summary>
        [StringLength(1)]
        [Display(Name = "")]
        public string Construction { get; set; }
        /// <summary></summary>
        [StringLength(5)]
        [Display(Name = "")]
        public string CDCodeNew { get; set; }
        /// <summary></summary>
        [StringLength(50)]
        [Display(Name = "")]
        public string FitType { get; set; }
        /// <summary></summary>
        [StringLength(50)]
        [Display(Name = "")]
        public string GearLine { get; set; }
        /// <summary></summary>
        [StringLength(5)]
        [Display(Name = "")]
        public string ThreadVersion { get; set; }

    }
}
