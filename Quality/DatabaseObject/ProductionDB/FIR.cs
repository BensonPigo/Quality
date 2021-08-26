using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class FIR
    {
        /// <summary>ID</summary>
        [Display(Name = "ID")]
        public Int64? ID { get; set; }

        /// <summary>採購單號</summary>
        [Display(Name = "採購單號")]
        public string POID { get; set; }

        /// <summary>大小項</summary>
        [Display(Name = "大小項")]
        public string SEQ1 { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string SEQ2 { get; set; }

        /// <summary>廠商代號</summary>
        [Display(Name = "廠商代號")]
        public string Suppid { get; set; }

        /// <summary>SCI Refno</summary>
        [Display(Name = "SCI Refno")]
        public string SCIRefno { get; set; }

        /// <summary>廠商Refno</summary>
        [Display(Name = "廠商Refno")]
        public string Refno { get; set; }

        /// <summary>收料單號</summary>
        [Display(Name = "收料單號")]
        public string ReceivingID { get; set; }

        /// <summary>補料單號</summary>
        [Display(Name = "補料單號")]
        public string ReplacementReportID { get; set; }

        /// <summary>收料數量</summary>
        [Display(Name = "收料數量")]
        public decimal ArriveQty { get; set; }

        /// <summary>總檢驗碼數</summary>
        [Display(Name = "總檢驗碼數")]
        public decimal TotalInspYds { get; set; }

        /// <summary>總檢驗點數</summary>
        [Display(Name = "總檢驗點數")]
        public decimal TotalDefectPoint { get; set; }

        /// <summary>檢驗結果</summary>
        [Display(Name = "檢驗結果")]
        public string Result { get; set; }

        /// <summary>備註</summary>
        [Display(Name = "備註")]
        public string Remark { get; set; }

        /// <summary>是否檢驗</summary>
        [Display(Name = "是否檢驗")]
        public bool Nonphysical { get; set; }

        /// <summary>布瑕疵點檢驗Result</summary>
        [Display(Name = "布瑕疵點檢驗Result")]
        public string Physical { get; set; }

        /// <summary>布瑕疵點檢驗Encode</summary>
        [Display(Name = "布瑕疵點檢驗Encode")]
        public bool PhysicalEncode { get; set; }

        /// <summary>布瑕疵點檢驗日期</summary>
        [Display(Name = "布瑕疵點檢驗日期")]
        public DateTime? PhysicalDate { get; set; }

        /// <summary>不需檢驗重量</summary>
        [Display(Name = "不需檢驗重量")]
        public bool nonWeight { get; set; }

        /// <summary>重量檢驗Result</summary>
        [Display(Name = "重量檢驗Result")]
        public string Weight { get; set; }

        /// <summary>重量檢驗Encode</summary>
        [Display(Name = "重量檢驗Encode")]
        public bool WeightEncode { get; set; }

        /// <summary>重量檢驗日期</summary>
        [Display(Name = "重量檢驗日期")]
        public DateTime? WeightDate { get; set; }

        /// <summary>不需檢驗 Shade Bond</summary>
        [Display(Name = "不需檢驗 Shade Bond")]
        public bool nonShadebond { get; set; }

        /// <summary>ShadeBond  Result</summary>
        [Display(Name = "ShadeBond  Result")]
        public string ShadeBond { get; set; }

        /// <summary>ShadeBond Encode</summary>
        [Display(Name = "ShadeBond Encode")]
        public bool ShadebondEncode { get; set; }

        /// <summary>ShadeBond Date</summary>
        [Display(Name = "ShadeBond Date")]
        public DateTime? ShadeBondDate { get; set; }

        /// <summary>不需檢驗漸進色</summary>
        [Display(Name = "不需檢驗漸進色")]
        public bool nonContinuity { get; set; }

        /// <summary>漸進色Result</summary>
        [Display(Name = "漸進色Result")]
        public string Continuity { get; set; }

        /// <summary>漸進色 Encode</summary>
        [Display(Name = "漸進色 Encode")]
        public bool ContinuityEncode { get; set; }

        /// <summary>漸進色日期</summary>
        [Display(Name = "漸進色日期")]
        public DateTime? ContinuityDate { get; set; }

        /// <summary>檢驗截止日</summary>
        [Display(Name = "檢驗截止日")]
        public DateTime? InspDeadline { get; set; }

        /// <summary>Approve 人員</summary>
        [Display(Name = "Approve 人員")]
        public string Approve { get; set; }

        /// <summary>Approve 時間</summary>
        [Display(Name = "Approve 時間")]
        public DateTime? ApproveDate { get; set; }

        /// <summary>新增人員</summary>
        [Display(Name = "新增人員")]
        public string AddName { get; set; }

        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }

        /// <summary>最後修改人員</summary>
        [Display(Name = "最後修改人員")]
        public string EditName { get; set; }

        /// <summary>最後修改時間</summary>
        [Display(Name = "最後修改時間")]
        public DateTime? EditDate { get; set; }

        /// <summary>狀態</summary>
        [Display(Name = "狀態")]
        public string Status { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string OldFabricUkey { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string OldFabricVer { get; set; }

        /// <summary></summary>
        public bool nonOdor { get; set; }

        /// <summary></summary>
        public string Odor { get; set; }

        /// <summary></summary>
        public bool OdorEncode { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? OdorDate { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string PhysicalInspector { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string WeightInspector { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string ShadeboneInspector { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string ContinuityInspector { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string OdorInspector { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string Moisture { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public bool NonMoisture { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? MoistureDate { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string MaterialCompositionGrp { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string MaterialCompositionItem { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public string MoistureStandardDesc { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public decimal MoistureStandard1 { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public Byte MoistureStandard1_Comparison { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public decimal MoistureStandard2 { get; set; }

        /// <summary></summary>
        [Display(Name = "")]
        public Byte MoistureStandard2_Comparison { get; set; }

    }
}
