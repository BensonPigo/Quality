using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class Orders
    {
        /// <summary>訂單單號</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "訂單單號")]
        public string ID { get; set; }
        /// <summary>品牌</summary>
        [StringLength(8)]
        [Display(Name = "品牌")]
        public string BrandID { get; set; }
        /// <summary>客戶品牌</summary>
        [StringLength(12)]
        [Display(Name = "客戶品牌")]
        public string ProgramID { get; set; }
        /// <summary>款式</summary>
        [StringLength(15)]
        [Display(Name = "款式")]
        public string StyleID { get; set; }
        /// <summary>季節</summary>
        [StringLength(10)]
        [Display(Name = "季節")]
        public string SeasonID { get; set; }
        /// <summary>專案代號</summary>
        [StringLength(5)]
        [Display(Name = "專案代號")]
        public string ProjectID { get; set; }
        /// <summary>訂單分類</summary>
        [StringLength(1)]
        [Display(Name = "訂單分類")]
        public string Category { get; set; }
        /// <summary>分類細項</summary>
        [StringLength(20)]
        [Display(Name = "分類細項")]
        public string OrderTypeID { get; set; }
        /// <summary>月份</summary>
        [StringLength(16)]
        [Display(Name = "月份")]
        public string BuyMonth { get; set; }
        /// <summary>進口國別</summary>
        [StringLength(2)]
        [Display(Name = "進口國別")]
        public string Dest { get; set; }
        /// <summary>場域模式</summary>
        [StringLength(25)]
        [Display(Name = "場域模式")]
        public string Model { get; set; }
        /// <summary>中國海關HS編碼</summary>
        [StringLength(14)]
        [Display(Name = "中國海關HS編碼")]
        public string HsCode1 { get; set; }
        /// <summary>中國海關HS編碼</summary>
        [StringLength(14)]
        [Display(Name = "中國海關HS編碼")]
        public string HsCode2 { get; set; }
        /// <summary>付款方式</summary>
        [StringLength(10)]
        [Display(Name = "付款方式")]
        public string PayTermARID { get; set; }
        /// <summary>交貨條件</summary>
        [StringLength(5)]
        [Display(Name = "交貨條件")]
        public string ShipTermID { get; set; }
        /// <summary>交貨方式</summary>
        [StringLength(30)]
        [Display(Name = "交貨方式")]
        public string ShipModeList { get; set; }
        /// <summary>CD#</summary>
        [StringLength(6)]
        [Display(Name = "CD#")]
        public string CdCodeID { get; set; }
        /// <summary>單件耗用產能</summary>
        [Display(Name = "單件耗用產能")]
        public decimal CPU { get; set; }
        /// <summary>訂單數量</summary>
        [Display(Name = "訂單數量")]
        public int? Qty { get; set; }
        /// <summary>款式單位</summary>
        [StringLength(8)]
        [Display(Name = "款式單位")]
        public string StyleUnit { get; set; }
        /// <summary>平均單價</summary>
        [Display(Name = "平均單價")]
        public decimal PoPrice { get; set; }
        /// <summary>確認單價</summary>
        [Display(Name = "確認單價")]
        public decimal CFMPrice { get; set; }
        /// <summary>幣別</summary>
        [StringLength(3)]
        [Display(Name = "幣別")]
        public string CurrencyID { get; set; }
        /// <summary>應付佣金%</summary>
        [Display(Name = "應付佣金%")]
        public decimal Commission { get; set; }
        /// <summary>工廠</summary>
        [StringLength(8)]
        [Display(Name = "工廠")]
        public string FactoryID { get; set; }
        /// <summary>客人的區域代號</summary>
        [StringLength(10)]
        [Display(Name = "客人的區域代號")]
        public string BrandAreaCode { get; set; }
        /// <summary>客人的工廠代碼</summary>
        [StringLength(10)]
        [Display(Name = "客人的工廠代碼")]
        public string BrandFTYCode { get; set; }
        /// <summary>每箱的包裝數量</summary>
        [Display(Name = "每箱的包裝數量")]
        public Int16 CTNQty { get; set; }

        /// <summary>客戶資料</summary>
        [StringLength(16)]
        [Display(Name = "客戶資料")]
        public string CustCDID { get; set; }
        /// <summary>客戶訂單單號</summary>
        [StringLength(30)]
        [Display(Name = "客戶訂單單號")]
        public string CustPONo { get; set; }
        /// <summary>客人的自訂欄位</summary>
        [StringLength(30)]
        [Display(Name = "客人的自訂欄位")]
        public string Customize1 { get; set; }
        /// <summary>客人的自訂欄位</summary>
        [StringLength(30)]
        [Display(Name = "客人的自訂欄位")]
        public string Customize2 { get; set; }
        /// <summary>客人的自訂欄位</summary>
        [StringLength(30)]
        [Display(Name = "客人的自訂欄位")]
        public string Customize3 { get; set; }
        /// <summary>訂單日期</summary>
        [Display(Name = "訂單日期")]
        public DateTime? CFMDate { get; set; }
        /// <summary>客戶交期</summary>
        [Display(Name = "客戶交期")]
        public DateTime? BuyerDelivery { get; set; }
        /// <summary>飛雁交期</summary>
        [Display(Name = "飛雁交期")]
        public DateTime? SciDelivery { get; set; }
        /// <summary>工廠上線日</summary>
        [Display(Name = "工廠上線日")]
        public DateTime? SewInLine { get; set; }
        /// <summary>工廠下線日</summary>
        [Display(Name = "工廠下線日")]
        public DateTime? SewOffLine { get; set; }
        /// <summary>裁剪上線日　</summary>
        [Display(Name = "裁剪上線日　")]
        public DateTime? CutInLine { get; set; }
        /// <summary>裁剪下線日</summary>
        [Display(Name = "裁剪下線日")]
        public DateTime? CutOffLine { get; set; }
        /// <summary>工廠出口日</summary>
        [Display(Name = "工廠出口日")]
        public DateTime? PulloutDate { get; set; }
        /// <summary>cmp單位</summary>
        [StringLength(8)]
        [Display(Name = "cmp單位")]
        public string CMPUnit { get; set; }
        /// <summary>cmp單價</summary>
        [Display(Name = "cmp單價")]
        public decimal CMPPrice { get; set; }
        /// <summary>CMPQ的確認日期</summary>
        [Display(Name = "CMPQ的確認日期")]
        public DateTime? CMPQDate { get; set; }
        /// <summary>CMPQ上的備註</summary>
        [Display(Name = "CMPQ上的備註")]
        public string CMPQRemark { get; set; }
        /// <summary>Each-con 確認日期</summary>
        [Display(Name = "Each-con 確認日期")]
        public DateTime? EachConsApv { get; set; }
        /// <summary>製造單確認日期</summary>
        [Display(Name = "製造單確認日期")]
        public DateTime? MnorderApv { get; set; }
        /// <summary>CRD date.</summary>
        [Display(Name = "CRD date.")]
        public DateTime? CRDDate { get; set; }
        /// <summary>Initial Plan Date</summary>
        [Display(Name = "Initial Plan Date")]
        public DateTime? InitialPlanDate { get; set; }
        /// <summary>Plan Date</summary>
        [Display(Name = "Plan Date")]
        public DateTime? PlanDate { get; set; }
        /// <summary>First production date</summary>
        [Display(Name = "First production date")]
        public DateTime? FirstProduction { get; set; }
        /// <summary>1st Production Lock</summary>
        [Display(Name = "1st Production Lock")]
        public DateTime? FirstProductionLock { get; set; }
        /// <summary>Orig. buyer delivery date:</summary>
        [Display(Name = "Orig. buyer delivery date:")]
        public DateTime? OrigBuyerDelivery { get; set; }
        /// <summary>ex-Country Date</summary>
        [Display(Name = "ex-Country Date")]
        public DateTime? ExCountry { get; set; }
        /// <summary>In DC Date</summary>
        [Display(Name = "In DC Date")]
        public DateTime? InDCDate { get; set; }
        /// <summary>Confirm Shipment Date</summary>
        [Display(Name = "Confirm Shipment Date")]
        public DateTime? CFMShipment { get; set; }
        /// <summary>P/F ETA</summary>
        [Display(Name = "P/F ETA")]
        public DateTime? PFETA { get; set; }
        /// <summary>包材的預計到貨日</summary>
        [Display(Name = "包材的預計到貨日")]
        public DateTime? PackLETA { get; set; }
        /// <summary>14天LOCK的採購單交期</summary>
        [Display(Name = "14天LOCK的採購單交期")]
        public DateTime? LETA { get; set; }
        /// <summary>訂單Handle</summary>
        [StringLength(10)]
        [Display(Name = "訂單Handle")]
        public string MRHandle { get; set; }
        /// <summary>組長</summary>
        [StringLength(10)]
        [Display(Name = "組長")]
        public string SMR { get; set; }
        /// <summary>Scan and Pack</summary>
        [Display(Name = "Scan and Pack")]
        public bool ScanAndPack { get; set; }
        /// <summary>VAS/SHAS</summary>
        [Display(Name = "VAS/SHAS")]
        public bool VasShas { get; set; }
        /// <summary>Special customer</summary>
        [Display(Name = "Special customer")]
        public bool SpecialCust { get; set; }
        /// <summary>棉紙</summary>
        [Display(Name = "棉紙")]
        public bool TissuePaper { get; set; }
        /// <summary>取消</summary>
        [Display(Name = "取消")]
        public bool Junk { get; set; }
        /// <summary>包裝說明</summary>
        [Display(Name = "包裝說明")]
        public string Packing { get; set; }
        /// <summary>大貨嘜頭(正面)</summary>
        [Display(Name = "大貨嘜頭(正面)")]
        public string MarkFront { get; set; }
        /// <summary>大貨嘜頭(背面)</summary>
        [Display(Name = "大貨嘜頭(背面)")]
        public string MarkBack { get; set; }
        /// <summary>大貨嘜頭(左面)</summary>
        [Display(Name = "大貨嘜頭(左面)")]
        public string MarkLeft { get; set; }
        /// <summary>大貨嘜頭(右面)</summary>
        [Display(Name = "大貨嘜頭(右面)")]
        public string MarkRight { get; set; }
        /// <summary>圖片與商標位置</summary>
        [Display(Name = "圖片與商標位置")]
        public string Label { get; set; }
        /// <summary>訂單備註</summary>
        [Display(Name = "訂單備註")]
        public string OrderRemark { get; set; }
        /// <summary>加工的展開方式</summary>
        [StringLength(1)]
        [Display(Name = "加工的展開方式")]
        public string ArtWorkCost { get; set; }
        /// <summary>標準成本</summary>
        [Display(Name = "標準成本")]
        public decimal StdCost { get; set; }
        /// <summary>包裝配比方式</summary>
        [StringLength(1)]
        [Display(Name = "包裝配比方式")]
        public string CtnType { get; set; }
        /// <summary>免費數量</summary>
        [Display(Name = "免費數量")]
        public int? FOCQty { get; set; }
        /// <summary>SMNotice Approved</summary>
        [Display(Name = "SMNotice Approved")]
        public DateTime? SMnorderApv { get; set; }
        /// <summary>免費</summary>
        [Display(Name = "免費")]
        public bool FOC { get; set; }
        /// <summary>mnoti_apv第二階段</summary>
        [Display(Name = "mnoti_apv第二階段")]
        public DateTime? MnorderApv2 { get; set; }
        /// <summary>Packing第二階段資料</summary>
        [Display(Name = "Packing第二階段資料")]
        public string Packing2 { get; set; }
        /// <summary>SAMPLE REASON</summary>
        [StringLength(5)]
        [Display(Name = "SAMPLE REASON")]
        public string SampleReason { get; set; }
        /// <summary>水洗測式</summary>
        [Display(Name = "水洗測式")]
        public bool RainwearTestPassed { get; set; }
        /// <summary>尺寸範圍</summary>
        [Display(Name = "尺寸範圍")]
        public string SizeRange { get; set; }
        /// <summary>物料出清</summary>
        [Display(Name = "物料出清")]
        public bool MTLComplete { get; set; }
        /// <summary>Special Mark</summary>
        [StringLength(5)]
        [Display(Name = "Special Mark")]
        public string SpecialMark { get; set; }
        /// <summary>延出備註</summary>
        [Display(Name = "延出備註")]
        public string OutstandingRemark { get; set; }
        /// <summary>延出原因修改人</summary>
        [StringLength(10)]
        [Display(Name = "延出原因修改人")]
        public string OutstandingInCharge { get; set; }
        /// <summary>延出原因修改時間</summary>
        [Display(Name = "延出原因修改時間")]
        public DateTime? OutstandingDate { get; set; }
        /// <summary>延出備註</summary>
        [StringLength(5)]
        [Display(Name = "延出備註")]
        public string OutstandingReason { get; set; }
        /// <summary>款式的唯一值</summary>
        [Display(Name = "款式的唯一值")]
        public Int64 StyleUkey { get; set; }
        /// <summary>採購單號</summary>
        [StringLength(13)]
        [Display(Name = "採購單號")]
        public string POID { get; set; }
        /// <summary></summary>
        [StringLength(13)]
        [Display(Name = "")]
        public string OrderComboID { get; set; }
        /// <summary>是否為非格子布或非Repeat或非Body Mapping</summary>
        [Display(Name = "是否為非格子布或非Repeat或非Body Mapping")]
        public bool IsNotRepeatOrMapping { get; set; }
        /// <summary>拆單的原訂單單號</summary>
        [StringLength(13)]
        [Display(Name = "拆單的原訂單單號")]
        public string SplitOrderId { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? FtyKPI { get; set; }
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
        /// <summary>車縫組別</summary>
        [StringLength(60)]
        [Display(Name = "車縫組別")]
        public string SewLine { get; set; }
        /// <summary>實際出貨日</summary>
        [Display(Name = "實際出貨日")]
        public DateTime? ActPulloutDate { get; set; }
        /// <summary>Production Schedules Remark</summary>
        [StringLength(100)]
        [Display(Name = "Production Schedules Remark")]
        public string ProdSchdRemark { get; set; }
        /// <summary>預估單</summary>
        [Display(Name = "預估單")]
        public bool IsForecast { get; set; }
        /// <summary>工廠自行接單</summary>
        [Display(Name = "工廠自行接單")]
        public bool LocalOrder { get; set; }
        /// <summary>Garment 結單</summary>
        [Display(Name = "Garment 結單")]
        public DateTime? GMTClose { get; set; }
        /// <summary>訂單總箱數</summary>
        [Display(Name = "訂單總箱數")]
        public int? TotalCTN { get; set; }
        /// <summary>cLog 已收到箱數</summary>
        [Display(Name = "cLog 已收到箱數")]
        public int? ClogCTN { get; set; }
        /// <summary>工廠已完成箱數</summary>
        [Display(Name = "工廠已完成箱數")]
        public int? FtyCTN { get; set; }
        /// <summary>大貨出貨結清</summary>
        [Display(Name = "大貨出貨結清")]
        public bool PulloutComplete { get; set; }
        /// <summary>大貨Ready 的日期</summary>
        [Display(Name = "大貨Ready 的日期")]
        public DateTime? ReadyDate { get; set; }
        /// <summary>Pullout箱數</summary>
        [Display(Name = "Pullout箱數")]
        public int? PulloutCTNQty { get; set; }
        /// <summary>Shipment(Pull out)完成</summary>
        [Display(Name = "Shipment(Pull out)完成")]
        public bool Finished { get; set; }
        /// <summary>Pull forward 訂單</summary>
        [Display(Name = "Pull forward 訂單")]
        public bool PFOrder { get; set; }
        /// <summary>工廠交期</summary>
        [Display(Name = "工廠交期")]
        public DateTime? SDPDate { get; set; }
        /// <summary>成衣首次通過檢驗日</summary>
        [Display(Name = "成衣首次通過檢驗日")]
        public DateTime? InspDate { get; set; }
        /// <summary>CFA檢驗結果</summary>
        [StringLength(1)]
        [Display(Name = "CFA檢驗結果")]
        public string InspResult { get; set; }
        /// <summary>CFA Finial檢驗人員</summary>
        [StringLength(10)]
        [Display(Name = "CFA Finial檢驗人員")]
        public string InspHandle { get; set; }
        /// <summary>KPI L/ETA</summary>
        [Display(Name = "KPI L/ETA")]
        public DateTime? KPILETA { get; set; }
        /// <summary>物料到達日</summary>
        [Display(Name = "物料到達日")]
        public DateTime? MTLETA { get; set; }
        /// <summary>Sewing Mtl ETA</summary>
        [Display(Name = "Sewing Mtl ETA")]
        public DateTime? SewETA { get; set; }
        /// <summary>Packing Mtl ETA</summary>
        [Display(Name = "Packing Mtl ETA")]
        public DateTime? PackETA { get; set; }
        /// <summary>物料出貨結清</summary>
        [StringLength(2)]
        [Display(Name = "物料出貨結清")]
        public string MTLExport { get; set; }
        /// <summary>FormA</summary>
        [StringLength(8)]
        [Display(Name = "FormA")]
        public string DoxType { get; set; }
        /// <summary>Factory Group</summary>
        [StringLength(8)]
        [Display(Name = "Factory Group")]
        public string FtyGroup { get; set; }
        /// <summary>Manufacturing Division ID</summary>
        [StringLength(8)]
        [Display(Name = "Manufacturing Division ID")]
        public string MDivisionID { get; set; }
        /// <summary>Cutting Ready Date</summary>
        [Display(Name = "Cutting Ready Date")]
        public DateTime? CutReadyDate { get; set; }
        /// <summary>Sewing Remark</summary>
        [StringLength(60)]
        [Display(Name = "Sewing Remark")]
        public string SewRemark { get; set; }
        /// <summary>倉庫關單</summary>
        [Display(Name = "倉庫關單")]
        public DateTime? WhseClose { get; set; }
        /// <summary>Subcon In From Sister Factory</summary>
        [Display(Name = "Subcon In From Sister Factory")]
        public bool SubconInSisterFty { get; set; }
        /// <summary>MC Handle</summary>
        [StringLength(10)]
        [Display(Name = "MC Handle")]
        public string MCHandle { get; set; }
        /// <summary>Local MR</summary>
        [StringLength(10)]
        [Display(Name = "Local MR")]
        public string LocalMR { get; set; }
        /// <summary>KPI Date變更理由</summary>
        [StringLength(5)]
        [Display(Name = "KPI Date變更理由")]
        public string KPIChangeReason { get; set; }
        /// <summary>MD Finished</summary>
        [Display(Name = "MD Finished")]
        public DateTime? MDClose { get; set; }
        /// <summary>MD 最後編輯人員</summary>
        [StringLength(10)]
        [Display(Name = "MD 最後編輯人員")]
        public string MDEditName { get; set; }
        /// <summary>MD 最後編輯時間</summary>
        [Display(Name = "MD 最後編輯時間")]
        public DateTime? MDEditDate { get; set; }
        /// <summary>最後收箱日期</summary>
        [Display(Name = "最後收箱日期")]
        public DateTime? ClogLastReceiveDate { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal CPUFactor { get; set; }
        /// <summary></summary>
        [StringLength(8)]
        [Display(Name = "")]
        public string SizeUnit { get; set; }
        /// <summary></summary>
        [StringLength(13)]
        [Display(Name = "")]
        public string CuttingSP { get; set; }
        /// <summary>是否為MixMarker </summary>
        [Display(Name = "是否為MixMarker ")]
        public int? IsMixMarker { get; set; }
        /// <summary></summary>
        [StringLength(1)]
        [Display(Name = "")]
        public string EachConsSource { get; set; }
        /// <summary>Each Cons. KPI Date (PMS only)</summary>
        [Display(Name = "Each Cons. KPI Date (PMS only)")]
        public DateTime? KPIEachConsApprove { get; set; }
        /// <summary>Cmpq KPI Date (PMS only)</summary>
        [Display(Name = "Cmpq KPI Date (PMS only)")]
        public DateTime? KPICmpq { get; set; }
        /// <summary>M Notice KPI Date (PMS only)</summary>
        [Display(Name = "M Notice KPI Date (PMS only)")]
        public DateTime? KPIMNotice { get; set; }
        /// <summary>Garment Complete ( From Trade)</summary>
        [StringLength(1)]
        [Display(Name = "Garment Complete ( From Trade)")]
        public string GMTComplete { get; set; }
        /// <summary>Global Foundation Range</summary>
        [Display(Name = "Global Foundation Range")]
        public bool GFR { get; set; }
        /// <summary></summary>
        public int? CfaCTN { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public int? DRYCTN { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public int? PackErrCTN { get; set; }
        /// <summary></summary>
        [StringLength(1)]
        [Display(Name = "")]
        public string ForecastSampleGroup { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal DyeingLoss { get; set; }
        /// <summary></summary>
        [StringLength(1)]
        [Display(Name = "")]
        public string SubconInType { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? LastProductionDate { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? EstPODD { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool AirFreightByBrand { get; set; }
        /// <summary></summary>
        [StringLength(13)]
        [Display(Name = "")]
        public string AllowanceComboID { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? ChangeMemoDate { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string BuyBack { get; set; }
        /// <summary></summary>
        [StringLength(13)]
        [Display(Name = "")]
        public string BuyBackOrderID { get; set; }
        /// <summary></summary>
        [StringLength(1)]
        public string ForecastCategory { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool OnSiteSample { get; set; }
        /// <summary>PulloutComplete 最後的更新時間</summary>
        [Display(Name = "PulloutComplete 最後的更新時間")]
        public DateTime? PulloutCmplDate { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool NeedProduction { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool IsBuyBack { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool KeepPanels { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "")]
        public string BuyBackReason { get; set; }
        /// <summary></summary>
        [Required]
        public bool IsBuyBackCrossArticle { get; set; }
        /// <summary></summary>
        [Required]
        public bool IsBuyBackCrossSizeCode { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? KpiEachConsCheck { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool NonRevenue { get; set; }
        /// <summary>Nike - Mercury - CAB</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "Nike - Mercury - CAB")]
        public string CAB { get; set; }
        /// <summary>Nike - Mercury - FinalDest</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "Nike - Mercury - FinalDest")]
        public string FinalDest { get; set; }
        /// <summary>Nike - Mercury - Customer_PO</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "Nike - Mercury - Customer_PO")]
        public string Customer_PO { get; set; }
        /// <summary>Nike - Mercury - AFS_STOCK_CATEGORY</summary>
        [Required]
        [StringLength(50)]
        [Display(Name = "Nike - Mercury - AFS_STOCK_CATEGORY")]
        public string AFS_STOCK_CATEGORY { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? CMPLTDATE { get; set; }
        /// <summary></summary>
        [StringLength(4)]
        [Display(Name = "")]
        public string DelayCode { get; set; }
        /// <summary></summary>
        [StringLength(100)]
        [Display(Name = "")]
        public string DelayDesc { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool HangerPack { get; set; }
        /// <summary></summary>
        [StringLength(5)]
        [Display(Name = "")]
        public string CDCodeNew { get; set; }
        /// <summary></summary>
        [StringLength(8)]
        [Display(Name = "")]
        public string SizeUnitWeight { get; set; }

    }
}
