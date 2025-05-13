using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ManufacturingExecutionDB
{
    public enum FinalInspectionSStepAction
    {
        Next,
        Previous,
        Current

    }

    public class FinalInspectionBasicStep
    {
        public long Ukey { get; set; }
        public string StepName { get; set; }
    }
    public class FinalInspectionBasicGeneral
    {
        public long Ukey { get; set; }
        public string ItemName { get; set; }
        public string ShowText { get; set; }
        public string GeneralColName { get; set; }
    }
    public class FinalInspectionBasicCheckList
    {
        public long Ukey { get; set; }
        public string Type { get; set; }
        public string ItemName { get; set; }
        public string CheckListColName { get; set; }
    }
    public class FinalInspection_Step
    {
        public string FinalInspectionID { get; set; }
        public string BrandID { get; set; }
        public int Seq { get; set; }
        public string StepName { get; set; }
        public long StepUkey { get; set; }
    }
    public class FinalInspectionBasicBrand_General
    {
        public long ID { get; set; }
        public long BasicGeneralUkey { get; set; }
        public string BrandID { get; set; }
    }
    public class FinalInspectionBasicBrand_CheckList
    {
        public long ID { get; set; }
        public long BasicCheckListUkey { get; set; }
        public string BrandID { get; set; }
    }

    public class FinalInspectionGeneral
    {
        public long ID { get; set; }
        public string FinalInspectionID { get; set; }
        public bool IsMaterialApproval { get; set; }
        public bool IsSealingSample { get; set; }
        public bool IsMetalDetection { get; set; }
        public bool IsFGWT { get; set; }
        public bool IsFGPT { get; set; }
        public bool IsTopSample { get; set; }
        public bool Is3rdPartyTestReport { get; set; }
        public bool IsPPSample { get; set; }
        public bool IsGBTestForChina { get; set; }
        public bool IsCPSIAForYounthStytle { get; set; }
        public bool IsQRSSample { get; set; }
        public bool IsFactoryDisclaimer { get; set; }
        public bool IsA01Compliance { get; set; }
        public bool IsCPSIACompliance { get; set; }
        public bool IsCustomerCountrySpecificCompliance { get; set; }

    }


    public class FinalInspectionCheckList
    {
        public long ID { get; set; }
        public string FinalInspectionID { get; set; }
        public bool IsCloseShade { get; set; }
        public bool IsHandfeel { get; set; }
        public bool IsAppearance { get; set; }
        public bool IsPrintEmbDecorations { get; set; }
        public bool IsEmbellishmentPrint { get; set; }
        public bool IsEmbellishmentBonding { get; set; }
        public bool IsEmbellishmentHT { get; set; }
        public bool IsEmbellishmentEMB { get; set; }
        public bool IsFiberContent { get; set; }
        public bool IsCareInstructions { get; set; }
        public bool IsDecorativeLabel { get; set; }
        public bool IsAdicomLabel { get; set; }
        public bool IsCountryofOrigion { get; set; }
        public bool IsSizeKey { get; set; }
        public bool Is8FlagLabel { get; set; }
        public bool IsAdditionalLabel { get; set; }
        public bool IsIdLabel { get; set; }
        public bool IsMainLabel { get; set; }
        public bool IsSizeLabel { get; set; }
        public bool IsCareContentLabel { get; set; }
        public bool IsBrandLabel { get; set; }
        public bool IsBlueSignLabel { get; set; }
        public bool IsLotLabel { get; set; }
        public bool IsSecurityLabel { get; set; }
        public bool IsSpecialLabel { get; set; }
        public bool IsVIDLabel { get; set; }
        public bool IsCNC { get; set; }
        public bool IsWovenlabel { get; set; }
        public bool IsTSize { get; set; }
        public bool IsCCLayout { get; set; }
        public bool IsShippingMark { get; set; }
        public bool IsPolytagMarking { get; set; }
        public bool IsColorSizeQty { get; set; }
        public bool IsHangtag { get; set; }
        public bool IsJokerTag { get; set; }
        public bool IsWWMT { get; set; }
        public bool IsChinaCIT { get; set; }
        public bool IsPolybagSticker { get; set; }
        public bool IsUCCSticker { get; set; }
        public bool IsPESheetMicropak { get; set; }
        public bool IsAdditionalHantage { get; set; }
        public bool IsUPCStickierHantage { get; set; }
        public bool IsGS1128Label { get; set; }
        public bool IsSecuritytag { get; set; }
        public bool IsPadPrintSizeLabel { get; set; }
    }

    public class FinalInspectionMoistureStandard
    {
        public string BrandID { get; set; }
        public string Category { get; set; }
        public decimal Standard { get; set; }
    }
}
