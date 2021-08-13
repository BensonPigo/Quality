using System;
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
    public class FinalInspection
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
        public string POID { get; set; }

        /// <summary>檢驗階段</summary>
        [Required]
        [StringLength(10)]
        [Display(Name = "檢驗階段")]
        public string InspectionStage { get; set; }

        /// <summary>檢驗次數</summary>
        [Required]
        [Display(Name = "檢驗次數")]
        public string InspectionTimes { get; set; }

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

        /// <summary>AcceptableQualityLevelsUkey</summary>
        [Display(Name = "AcceptableQualityLevelsUkey")]
        public string AcceptableQualityLevelsUkey { get; set; }

        /// <summary>樣本數量</summary>
        [Display(Name = "樣本數量")]
        public int? SampleSize { get; set; }

        /// <summary>允許檢驗失敗數量</summary>
        [Display(Name = "允許檢驗失敗數量")]
        public int? AcceptQty { get; set; }

        /// <summary>是否收到Fabric Approval文件</summary>
        [Display(Name = "是否收到Fabric Approval文件")]
        public string FabricApprovalDoc { get; set; }

        /// <summary>是否收到Sealing Sample文件</summary>
        [Display(Name = "是否收到Sealing Sample文件")]
        public string SealingSampleDoc { get; set; }

        /// <summary>是否收到Metal Detection文件</summary>
        [Display(Name = "是否收到Metal Detection文件")]
        public string MetalDetectionDoc { get; set; }

        /// <summary>是否收到Garment Washing文件</summary>
        [Display(Name = "是否收到Garment Washing文件")]
        public string GarmentWashingDoc { get; set; }

        /// <summary>是否收到Close/ Shade</summary>
        [Display(Name = "是否收到Close/ Shade")]
        public string CheckCloseShade { get; set; }

        /// <summary>是否收到Handfeel</summary>
        [Display(Name = "是否收到Handfeel")]
        public string CheckHandfeel { get; set; }

        /// <summary>是否收到Appearance</summary>
        [Display(Name = "是否收到Appearance")]
        public string CheckAppearance { get; set; }

        /// <summary>是否收到Print/ Emb Decorations</summary>
        [Display(Name = "是否收到Print/ Emb Decorations")]
        public string CheckPrintEmbDecorations { get; set; }

        /// <summary>是否收到Fiber Content</summary>
        [Display(Name = "是否收到Fiber Content")]
        public string CheckFiberContent { get; set; }

        /// <summary>是否收到Care Instructions</summary>
        [Display(Name = "是否收到Care Instructions")]
        public string CheckCareInstructions { get; set; }

        /// <summary>是否收到Decorative Label</summary>
        [Display(Name = "是否收到Decorative Label")]
        public string CheckDecorativeLabel { get; set; }

        /// <summary>是否收到Adicom Label</summary>
        [Display(Name = "是否收到Adicom Label")]
        public string CheckAdicomLabel { get; set; }

        /// <summary>是否收到Country of Origion</summary>
        [Display(Name = "是否收到Country of Origion")]
        public string CheckCountryofOrigion { get; set; }

        /// <summary>是否收到Size Key</summary>
        [Display(Name = "是否收到Size Key")]
        public string CheckSizeKey { get; set; }

        /// <summary>是否收到8-Flag Label</summary>
        [Display(Name = "是否收到8-Flag Label")]
        public string Check8FlagLabel { get; set; }

        /// <summary>是否收到Additional Label</summary>
        [Display(Name = "是否收到Additional Label")]
        public string CheckAdditionalLabel { get; set; }

        /// <summary>是否收到Shipping Mark</summary>
        [Display(Name = "是否收到Shipping Mark")]
        public string CheckShippingMark { get; set; }

        /// <summary>是否收到Polytag/ Marketing</summary>
        [Display(Name = "是否收到Polytag/ Marketing")]
        public string CheckPolytagMarketing { get; set; }

        /// <summary>是否收到Color/ Size/ Qty</summary>
        [Display(Name = "是否收到Color/ Size/ Qty")]
        public string CheckColorSizeQty { get; set; }

        /// <summary>是否收到Hangtag</summary>
        [Display(Name = "是否收到Hangtag")]
        public string CheckHangtag { get; set; }

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

        /// <summary>完成比例</summary>
        [Display(Name = "完成比例")]
        public string ProductionStatus { get; set; }

        /// <summary>Pass/ Fail/ On-going</summary>
        [StringLength(8)]
        [Display(Name = "Pass/ Fail/ On-going")]
        public string InspectionResult { get; set; }

        /// <summary>Ship/ On Hold</summary>
        [StringLength(7)]
        [Display(Name = "Ship/ On Hold")]
        public string ShipmentStatus { get; set; }

        /// <summary>Others-備註</summary>
        [StringLength(-1)]
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

    }
}
