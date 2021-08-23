using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*Garment Test(GarmentTest_Detail) 詳細敘述如下*/
    /// <summary>
    /// Garment Test
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>
    public class GarmentTest_Detail
    {
        /// <summary>ID</summary>
        [Required]
        [Display(Name = "ID")]
        public Int64 ID { get; set; }

        /// <summary>檢驗序號</summary>
        [Required]
        [Display(Name = "檢驗序號")]
        public int? No { get; set; }
        /// <summary>結果</summary>
        [StringLength(1)]
        [Display(Name = "結果")]
        public string Result { get; set; }
        /// <summary>檢驗日期</summary>
        [Display(Name = "檢驗日期")]
        public DateTime? inspdate { get; set; }
        /// <summary>測試人員</summary>
        [StringLength(10)]
        [Display(Name = "測試人員")]
        public string inspector { get; set; }
        /// <summary>備註</summary>
        [StringLength(-1)]
        [Display(Name = "備註")]
        public string Remark { get; set; }
        /// <summary>寄出人員</summary>
        [StringLength(10)]
        [Display(Name = "寄出人員")]
        public string Sender { get; set; }
        /// <summary>寄出時間</summary>
        [Display(Name = "寄出時間")]
        public DateTime? SendDate { get; set; }
        /// <summary>收件人員</summary>
        [StringLength(10)]
        [Display(Name = "收件人員")]
        public string Receiver { get; set; }
        /// <summary>收件時間</summary>
        [Display(Name = "收件時間")]
        public DateTime? ReceiveDate { get; set; }
        /// <summary>新增者</summary>
        [StringLength(10)]
        [Display(Name = "新增者")]
        public string AddName { get; set; }
        /// <summary>新增時間</summary>
        [Display(Name = "新增時間")]
        public DateTime? AddDate { get; set; }
        /// <summary>編輯者</summary>
        [StringLength(10)]
        [Display(Name = "編輯者")]
        public string EditName { get; set; }
        /// <summary>編輯時間</summary>
        [Display(Name = "編輯時間")]
        public DateTime? EditDate { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string OldUkey { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public DateTime? SubmitDate { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public int? ArrivedQty { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool LineDry { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public int? Temperature { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool TumbleDry { get; set; }
        /// <summary></summary>
        [StringLength(10)]
        [Display(Name = "")]
        public string Machine { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool HandWash { get; set; }
        /// <summary></summary>
        [StringLength(50)]
        [Display(Name = "")]
        public string Composition { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool Neck { get; set; }
        /// <summary></summary>
        [StringLength(15)]
        [Display(Name = "")]
        public string Status { get; set; }
        /// <summary></summary>
        [StringLength(8)]
        [Display(Name = "")]
        public string SizeCode { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(100)]
        [Display(Name = "")]
        public string LOtoFactory { get; set; }
        /// <summary></summary>
        [Required]
        [StringLength(20)]
        [Display(Name = "")]
        public string MtlTypeID { get; set; }
        /// <summary>All basic Fabrics ≥ 50% natural fibres ; 布料 50% (含) 以上是天然纖維</summary>
        [Required]
        [Display(Name = "All basic Fabrics ≥ 50% natural fibres ; 布料 50% (含) 以上是天然纖維")]
        public bool Above50NaturalFibres { get; set; }
        /// <summary>All basic Fabrics ≥ 50% synthetic fibres (ex. polyester) ; 布料 50% (含)以上是合成纖維 (e.x. 聚酯纖維)</summary>
        [Required]
        [Display(Name = "All basic Fabrics ≥ 50% synthetic fibres (ex. polyester) ; 布料 50% (含)以上是合成纖維 (e.x. 聚酯纖維)")]
        public bool Above50SyntheticFibres { get; set; }
        /// <summary>測試的訂單號碼</summary>
        [Required]
        [StringLength(13)]
        [Display(Name = "測試的訂單號碼")]
        public string OrderID { get; set; }
        /// <summary>紀錄該次測試中是否包含 PHX-AP0450 SeamBrakage 的測試</summary>
        [Required]
        [Display(Name = "紀錄該次測試中是否包含 PHX-AP0450 SeamBrakage 的測試")]
        public string NonSeamBreakageTest { get; set; }
        /// <summary>Adidas-PHX-AP0450 SeamBrakage 檢驗結果</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "Adidas-PHX-AP0450 SeamBrakage 檢驗結果")]
        public string SeamBreakageResult { get; set; }
        /// <summary>Adidas-PHX-AP0451 檢驗結果</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "Adidas-PHX-AP0451 檢驗結果")]
        public string OdourResult { get; set; }
        /// <summary>Adidas-PHX-AP0701/PHX-AP0710 檢驗結果</summary>
        [Required]
        [StringLength(1)]
        [Display(Name = "Adidas-PHX-AP0701/PHX-AP0710 檢驗結果")]
        public string WashResult { get; set; }
        /// <summary>測試前的照片</summary>
        [Display(Name = "測試前的照片")]
        public Byte[] TestBeforePicture { get; set; }
        /// <summary>測試後的照片</summary>
        [Display(Name = "測試後的照片")]
        public Byte[] TestAfterPicture { get; set; }

    }
}
