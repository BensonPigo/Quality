using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    /*(FinalInspection) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/10 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/10   1.00    Admin        Create
    /// </history>
    public class FinalInspection : ResultModelBase<FinalInspection>
    {
        /// <summary>單號</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "單號")]
        public string ID { get; set; }

        /// <summary>訂單PO</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "訂單PO")]
        public string CustPONO { get; set; }
        public string BrandID { get; set; }
        public bool ReInspection { get; set; }

        /// <summary>檢驗階段</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "檢驗階段")]
        public string InspectionStage { get; set; }

        /// <summary>檢驗次數</summary>
        [Required]
        [Display(Name = "檢驗次數")]
        public int InspectionTimes { get; set; }

        /// <summary>廠別</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "廠別")]
        public string FactoryID { get; set; }

        /// <summary>M</summary>
        [Required]
        [StringLength(8)]
        [Display(Name = "M")]
        public string MDivisionID { get; set; }

        /// <summary>稽核日</summary>
        [Display(Name = "稽核日")]
        public DateTime? AuditDate { get; set; }

        /// <summary>產線</summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "產線")]
        public string SewingLineID { get; set; }
        public string Shift { get; set; }
        public string Team { get; set; }

        /// <summary>AcceptableQualityLevelsUkey</summary>
        [Display(Name = "AcceptableQualityLevelsUkey")]
        public long AcceptableQualityLevelsUkey { get; set; }
        public long AcceptableQualityLevelsProUkey { get; set; }

        /// <summary>樣本數量</summary>
        [Display(Name = "樣本數量")]
        public int? SampleSize { get; set; }

        /// <summary>允許檢驗失敗數量</summary>
        [Display(Name = "允許檢驗失敗數量")]
        public int? AcceptQty { get; set; }

        /// <summary>是否收到Fabric Approval文件</summary>
        [Display(Name = "是否收到Fabric Approval文件")]
        public bool FabricApprovalDoc { get; set; }

        /// <summary>是否收到Sealing Sample文件</summary>
        [Display(Name = "是否收到Sealing Sample文件")]
        public bool SealingSampleDoc { get; set; }

        /// <summary>是否收到Metal Detection文件</summary>
        [Display(Name = "是否收到Metal Detection文件")]
        public bool MetalDetectionDoc { get; set; }

        /// <summary>是否收到Garment Washing文件</summary>
        [Display(Name = "是否收到Garment Washing文件")]
        public bool GarmentWashingDoc { get; set; }

        /// <summary>是否收到Close/ Shade</summary>
        [Display(Name = "是否收到Close/ Shade")]
        public bool CheckCloseShade { get; set; }

        /// <summary>是否收到Handfeel</summary>
        [Display(Name = "是否收到Handfeel")]
        public bool CheckHandfeel { get; set; }

        /// <summary>是否收到Appearance</summary>
        [Display(Name = "是否收到Appearance")]
        public bool CheckAppearance { get; set; }

        /// <summary>是否收到Print/ Emb Decorations</summary>
        [Display(Name = "是否收到Print/ Emb Decorations")]
        public bool CheckPrintEmbDecorations { get; set; }

        /// <summary>是否收到Fiber Content</summary>
        [Display(Name = "是否收到Fiber Content")]
        public bool CheckFiberContent { get; set; }

        /// <summary>是否收到Care Instructions</summary>
        [Display(Name = "是否收到Care Instructions")]
        public bool CheckCareInstructions { get; set; }

        /// <summary>是否收到Decorative Label</summary>
        [Display(Name = "是否收到Decorative Label")]
        public bool CheckDecorativeLabel { get; set; }

        /// <summary>是否收到Adicom Label</summary>
        [Display(Name = "是否收到Adicom Label")]
        public bool CheckAdicomLabel { get; set; }

        /// <summary>是否收到Country of Origion</summary>
        [Display(Name = "是否收到Country of Origion")]
        public bool CheckCountryofOrigion { get; set; }

        /// <summary>是否收到Size Key</summary>
        [Display(Name = "是否收到Size Key")]
        public bool CheckSizeKey { get; set; }

        /// <summary>是否收到8-Flag Label</summary>
        [Display(Name = "是否收到8-Flag Label")]
        public bool Check8FlagLabel { get; set; }

        /// <summary>是否收到Additional Label</summary>
        [Display(Name = "是否收到Additional Label")]
        public bool CheckAdditionalLabel { get; set; }

        /// <summary>是否收到Shipping Mark</summary>
        [Display(Name = "是否收到Shipping Mark")]
        public bool CheckShippingMark { get; set; }

        /// <summary>是否收到Polybag/ Marking</summary>
        [Display(Name = "是否收到Polybag/ Marking")]
        public bool CheckPolytagMarking { get; set; }

        /// <summary>是否收到Color/ Size/ Qty</summary>
        [Display(Name = "是否收到Color/ Size/ Qty")]
        public bool CheckColorSizeQty { get; set; }

        /// <summary>是否收到Hangtag</summary>
        [Display(Name = "是否收到Hangtag")]
        public bool CheckHangtag { get; set; }

        /// <summary>檢驗成功數量</summary>
        [Display(Name = "檢驗成功數量")]
        public int? PassQty { get; set; }

        /// <summary>檢驗失敗數量</summary>
        [Display(Name = "檢驗失敗數量")]
        public int? RejectQty { get; set; }

        /// <summary>beautiful product audit的完美數量</summary>
        [Display(Name = "beautiful product audit的完美數量")]
        public int? BAQty { get; set; }

        /// <summary>檢驗人員</summary>
        [StringLength(10)]
        [Display(Name = "檢驗人員")]
        public string CFA { get; set; }
        public string Clerk { get; set; }

        /// <summary>完成比例</summary>
        [Display(Name = "完成比例")]
        public decimal? ProductionStatus { get; set; }

        /// <summary>Pass/ Fail/ On-going</summary>
        [StringLength(8)]
        [Display(Name = "Pass/ Fail/ On-going")]
        public string InspectionResult { get; set; }

        /// <summary>Ship/ On Hold</summary>
        [StringLength(7)]
        [Display(Name = "Ship/ On Hold")]
        public string ShipmentStatus { get; set; }

        /// <summary>Others-備註</summary>
        [Display(Name = "Others-備註")]
        public string OthersRemark { get; set; }

        /// <summary>提交檢驗日期</summary>
        [Display(Name = "提交檢驗日期")]
        public DateTime? SubmitDate { get; set; }

        /// <summary>InspectionStep</summary>
        [StringLength(16)]
        [Display(Name = "InspectionStep")]
        public string InspectionStep { get; set; }

        /// <summary>新建人員</summary>
        [StringLength(10)]
        [Display(Name = "新建人員")]
        public string AddName { get; set; }

        /// <summary>新建日期</summary>
        [Display(Name = "新建日期")]
        public DateTime? AddDate { get; set; }

        /// <summary>最後編輯人員</summary>
        [StringLength(10)]
        [Display(Name = "最後編輯人員")]
        public string EditName { get; set; }

        /// <summary>最後編輯日</summary>
        [Display(Name = "最後編輯日")]
        public DateTime? EditDate { get; set; }

        public bool HasOtherImage { get; set; }

        [Display(Name = "FGPT")]
        public bool CheckFGPT { get; set; }

        public string FGWT { get; set; }

        public string FGPT { get; set; }

        public bool ISFD { get; set; }

        public long MeasurementAQLUkey { get; set; }
        public int MeasurementSampleSize { get; set; }
        public int MeasurementAcceptQty { get; set; }
        public Dictionary<string, bool> GeneralDic
        {
            get 
            { 
                 Dictionary<string, bool> result = new Dictionary<string, bool>();
                if (this.finalInspectionGeneral != null)
                {
                    result.Add("IsMaterialApproval", this.finalInspectionGeneral.IsMaterialApproval);
                    result.Add("IsSealingSample", this.finalInspectionGeneral.IsSealingSample);
                    result.Add("IsMetalDetection", this.finalInspectionGeneral.IsMetalDetection);
                    result.Add("IsFGWT", this.finalInspectionGeneral.IsFGWT);
                    result.Add("IsFGPT", this.finalInspectionGeneral.IsFGPT);
                    result.Add("IsTopSample", this.finalInspectionGeneral.IsTopSample);
                    result.Add("Is3rdPartyTestReport", this.finalInspectionGeneral.Is3rdPartyTestReport);
                    result.Add("IsPPSample", this.finalInspectionGeneral.IsPPSample);
                    result.Add("IsGBTestForChina", this.finalInspectionGeneral.IsGBTestForChina);
                    result.Add("IsCPSIAForYounthStytle", this.finalInspectionGeneral.IsCPSIAForYounthStytle);
                    result.Add("IsQRSSample", this.finalInspectionGeneral.IsQRSSample);
                    result.Add("IsFactoryDisclaimer", this.finalInspectionGeneral.IsFactoryDisclaimer);
                    result.Add("IsA01Compliance", this.finalInspectionGeneral.IsA01Compliance);
                    result.Add("IsCPSIACompliance", this.finalInspectionGeneral.IsCPSIACompliance);
                    result.Add("IsCustomerCountrySpecificCompliance", this.finalInspectionGeneral.IsCustomerCountrySpecificCompliance);
                }

                return result;
            }
        }
        public Dictionary<string, bool> CheckListListDic
        {
            get
            {
                Dictionary<string, bool> result = new Dictionary<string, bool>();
                if (this.finalInspectionCheckList != null)
                {
                    result.Add("IsCloseShade", this.finalInspectionCheckList.IsCloseShade);
                    result.Add("IsHandfeel", this.finalInspectionCheckList.IsHandfeel);
                    result.Add("IsAppearance", this.finalInspectionCheckList.IsAppearance);
                    result.Add("IsPrintEmbDecorations", this.finalInspectionCheckList.IsPrintEmbDecorations);
                    result.Add("IsEmbellishmentPrint", this.finalInspectionCheckList.IsEmbellishmentPrint);
                    result.Add("IsEmbellishmentBonding", this.finalInspectionCheckList.IsEmbellishmentBonding);
                    result.Add("IsEmbellishmentHT", this.finalInspectionCheckList.IsEmbellishmentHT);
                    result.Add("IsEmbellishmentEMB", this.finalInspectionCheckList.IsEmbellishmentEMB);
                    result.Add("IsFiberContent", this.finalInspectionCheckList.IsFiberContent);
                    result.Add("IsCareInstructions", this.finalInspectionCheckList.IsCareInstructions);
                    result.Add("IsDecorativeLabel", this.finalInspectionCheckList.IsDecorativeLabel);
                    result.Add("IsAdicomLabel", this.finalInspectionCheckList.IsAdicomLabel);
                    result.Add("IsCountryofOrigion", this.finalInspectionCheckList.IsCountryofOrigion);
                    result.Add("IsSizeKey", this.finalInspectionCheckList.IsSizeKey);
                    result.Add("Is8FlagLabel", this.finalInspectionCheckList.Is8FlagLabel);
                    result.Add("IsAdditionalLabel", this.finalInspectionCheckList.IsAdditionalLabel);
                    result.Add("IsIdLabel", this.finalInspectionCheckList.IsIdLabel);
                    result.Add("IsMainLabel", this.finalInspectionCheckList.IsMainLabel);
                    result.Add("IsSizeLabel", this.finalInspectionCheckList.IsSizeLabel);
                    result.Add("IsCareContentLabel", this.finalInspectionCheckList.IsCareContentLabel);
                    result.Add("IsBrandLabel", this.finalInspectionCheckList.IsBrandLabel);
                    result.Add("IsBlueSignLabel", this.finalInspectionCheckList.IsBlueSignLabel);
                    result.Add("IsLotLabel", this.finalInspectionCheckList.IsLotLabel);
                    result.Add("IsSecurityLabel", this.finalInspectionCheckList.IsSecurityLabel);
                    result.Add("IsSpecialLabel", this.finalInspectionCheckList.IsSpecialLabel);
                    result.Add("IsVIDLabel", this.finalInspectionCheckList.IsVIDLabel);
                    result.Add("IsCNC", this.finalInspectionCheckList.IsCNC);
                    result.Add("IsWovenlabel", this.finalInspectionCheckList.IsWovenlabel);
                    result.Add("IsTSize", this.finalInspectionCheckList.IsTSize);
                    result.Add("IsCCLayout", this.finalInspectionCheckList.IsCCLayout);
                    result.Add("IsShippingMark", this.finalInspectionCheckList.IsShippingMark);
                    result.Add("IsPolytagMarking", this.finalInspectionCheckList.IsPolytagMarking);
                    result.Add("IsColorSizeQty", this.finalInspectionCheckList.IsColorSizeQty);
                    result.Add("IsHangtag", this.finalInspectionCheckList.IsHangtag);
                    result.Add("IsJokerTag", this.finalInspectionCheckList.IsJokerTag);
                    result.Add("IsWWMT", this.finalInspectionCheckList.IsWWMT);
                    result.Add("IsChinaCIT", this.finalInspectionCheckList.IsChinaCIT);
                    result.Add("IsPolybagSticker", this.finalInspectionCheckList.IsPolybagSticker);
                    result.Add("IsUCCSticker", this.finalInspectionCheckList.IsUCCSticker);
                    result.Add("IsPESheetMicropak", this.finalInspectionCheckList.IsPESheetMicropak);
                    result.Add("IsAdditionalHantage", this.finalInspectionCheckList.IsAdditionalHantage);
                    result.Add("IsUPCStickierHantage", this.finalInspectionCheckList.IsUPCStickierHantage);
                    result.Add("IsGS1128Label", this.finalInspectionCheckList.IsGS1128Label);

                }

                return result;
            }
        }
        public FinalInspectionGeneral finalInspectionGeneral { get; set; }
        public FinalInspectionCheckList finalInspectionCheckList { get; set; }
        public List<FinalInspectionBasicGeneral> GeneralList { get; set; }
        public List<FinalInspectionBasicCheckList> CheckListList { get; set; }
    }
}
