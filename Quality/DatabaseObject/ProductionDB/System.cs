using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*系統參數檔(System) 詳細敘述如下*/
    /// <summary>
    /// 系統參數檔
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/12 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/12   1.00    Admin        Create
    /// </history>
    public class System
    {
        /// <summary>Mail Server IP</summary>
        [StringLength(60)]
        [Display(Name = "Mail Server IP")]
        public string Mailserver { get; set; }
        /// <summary>寄件者名稱(有@)</summary>
        [StringLength(40)]
        [Display(Name = "寄件者名稱(有@)")]
        public string Sendfrom { get; set; }
        /// <summary>Email ID</summary>
        [StringLength(40)]
        [Display(Name = "Email ID")]
        public string EmailID { get; set; }
        /// <summary>Email Password</summary>
        [StringLength(20)]
        [Display(Name = "Email Password")]
        public string EmailPwd { get; set; }
        /// <summary>Picture Path</summary>
        [StringLength(80)]
        [Display(Name = "Picture Path")]
        public string PicPath { get; set; }
        /// <summary>CPU和TMS換算值(1400)</summary>
        [Required]
        [Display(Name = "CPU和TMS換算值(1400)")]
        public int? StdTMS { get; set; }
        /// <summary>夾檔存放位置</summary>
        [StringLength(80)]
        [Display(Name = "夾檔存放位置")]
        public string ClipPath { get; set; }
        /// <summary>FTP IP</summary>
        [StringLength(36)]
        [Display(Name = "FTP IP")]
        public string FtpIP { get; set; }
        /// <summary>FTP ID</summary>
        [StringLength(10)]
        [Display(Name = "FTP ID")]
        public string FtpID { get; set; }
        /// <summary>FTP PASSWORD</summary>
        [StringLength(36)]
        [Display(Name = "FTP PASSWORD")]
        public string FtpPwd { get; set; }
        /// <summary>車縫日報表月結日</summary>
        [Display(Name = "車縫日報表月結日")]
        public DateTime? SewLock { get; set; }
        /// <summary>銷樣單的倍數</summary>
        [Display(Name = "銷樣單的倍數")]
        public int SampleRate { get; set; }
        /// <summary>Pullout Report月結日</summary>
        [Display(Name = "Pullout Report月結日")]
        public DateTime? PullLock { get; set; }
        /// <summary>區域代號</summary>
        [Required]
        [StringLength(3)]
        [Display(Name = "區域代號")]
        public string RgCode { get; set; }
        /// <summary>資料上傳位置</summary>
        [StringLength(60)]
        [Display(Name = "資料上傳位置")]
        public string ImportDataPath { get; set; }
        /// <summary>資料上傳檔名</summary>
        [StringLength(60)]
        [Display(Name = "資料上傳檔名")]
        public string ImportDataFileName { get; set; }
        /// <summary>資料下載位置</summary>
        [StringLength(60)]
        [Display(Name = "資料下載位置")]
        public string ExportDataPath { get; set; }
        /// <summary>當地幣別</summary>
        [StringLength(4)]
        [Display(Name = "當地幣別")]
        public string CurrencyID { get; set; }
        /// <summary>美金匯率</summary>
        [Display(Name = "美金匯率")]
        public decimal USDRate { get; set; }
        /// <summary>Local Purchase Approve Name</summary>
        [StringLength(10)]
        [Display(Name = "Local Purchase Approve Name")]
        public string POApproveName { get; set; }
        /// <summary>Local Purchase 自動Approve的天數</summary>
        [Display(Name = "Local Purchase 自動Approve的天數")]
        public int POApproveDay { get; set; }
        /// <summary>裁剪上線日為Sewing Inline的前幾日</summary>
        [Display(Name = "裁剪上線日為Sewing Inline的前幾日")]
        public int CutDay { get; set; }
        /// <summary>帳號保留字</summary>
        [StringLength(1)]
        [Display(Name = "帳號保留字")]
        public string AccountKeyword { get; set; }
        /// <summary>Production Ready day</summary>
        [Display(Name = "Production Ready day")]
        public int ReadyDay { get; set; }
        /// <summary>VN出口報關調整營收倍數</summary>
        [Display(Name = "VN出口報關調整營收倍數")]
        public decimal VNMultiple { get; set; }
        /// <summary>Material Leat-time</summary>
        [Display(Name = "Material Leat-time")]
        public int MtlLeadTime { get; set; }
        /// <summary>Rate Type</summary>
        [StringLength(2)]
        [Display(Name = "Rate Type")]
        public string ExchangeID { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string RFIDServerName { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string RFIDDatabaseName { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string RFIDLoginId { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string RFIDLoginPwd { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string RFIDTable { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public decimal ProphetSingleSizeDeduct { get; set; }
        /// <summary></summary>
        [StringLength(8)]
        [Display(Name = "")]
        public string PrintingSuppID { get; set; }
        /// <summary>驗布機延遲時間(秒)</summary>
        [Display(Name = "驗布機延遲時間(秒)")]
        public decimal QCMachineDelayTime { get; set; }
        /// <summary></summary>
        [StringLength(15)]
        [Display(Name = "")]
        public string APSLoginId { get; set; }
        /// <summary></summary>
        [StringLength(15)]
        [Display(Name = "")]
        public string APSLoginPwd { get; set; }
        /// <summary></summary>
        [StringLength(130)]
        [Display(Name = "")]
        public string SQLServerName { get; set; }
        /// <summary></summary>
        [StringLength(15)]
        [Display(Name = "")]
        public string APSDatabaseName { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool RFIDMiddlewareInRFIDServer { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool UseAutoScanPack { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool MtlAutoLock { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool InspAutoLockAcc { get; set; }
        /// <summary></summary>
        [StringLength(80)]
        [Display(Name = "")]
        public string ShippingMarkPath { get; set; }
        /// <summary></summary>
        [StringLength(80)]
        [Display(Name = "")]
        public string StyleSketch { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string ARKServerName { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string ARKDatabaseName { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string ARKLoginId { get; set; }
        /// <summary></summary>
        [StringLength(20)]
        [Display(Name = "")]
        public string ARKLoginPwd { get; set; }
        /// <summary></summary>
        [StringLength(80)]
        [Display(Name = "")]
        public string MarkerInputPath { get; set; }
        /// <summary></summary>
        [StringLength(80)]
        [Display(Name = "")]
        public string MarkerOutputPath { get; set; }
        /// <summary></summary>
        [StringLength(80)]
        [Display(Name = "")]
        public string ReplacementReport { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool CuttingP10mustCutRef { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool Automation { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool AutomationAutoRunTime { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool CanReviseDailyLockData { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool AutoGenerateByTone { get; set; }
        /// <summary>Misc Local Purchase Approve Name</summary>
        [StringLength(10)]
        [Display(Name = "Misc Local Purchase Approve Name")]
        public string MiscPOApproveName { get; set; }
        /// <summary>Misc Local Purchase �۰�Approve���Ѽ�</summary>
        [Display(Name = "Misc Local Purchase ")]
        public int MiscPOApproveDay { get; set; }
        /// <summary>QMS</summary>
        [Required]
        [Display(Name = "")]
        public bool QMSAutoAdjustMtl { get; set; }
        /// <summary>���ҽd���ɸ�|</summary>
        [Required]
        [StringLength(80)]
        [Display(Name = "")]
        public string ShippingMarkTemplatePath { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool WIP_FollowCutOutput { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public bool NoRestrictOrdersDelivery { get; set; }
        /// <summary></summary>
        [Display(Name = "")]
        public bool WIP_ByShell { get; set; }
        /// <summary>RF Card Erase Before Printing</summary>
        [Required]
        [Display(Name = "RF Card Erase Before Printing")]
        public bool RFCardEraseBeforePrinting { get; set; }
        /// <summary>PH ���u�����@�Ѫ� CPU</summary>
        [Display(Name = "")]
        public int? SewlineAvgCPU { get; set; }
        /// <summary>�H�U�P�_ smalll logo �����׼з� (CM)</summary>
        [Display(Name = "")]
        public decimal? SmallLogoCM { get; set; }
        /// <summary>�O�_�S�LWS���ˬdRFID�d���ƨϥ�</summary>
        [Required]
        [Display(Name = "")]
        public bool CheckRFIDCardDuplicateByWebservice { get; set; }
        /// <summary>�O�_combine subprocess����bundle</summary>
        [Required]
        [Display(Name = "")]
        public bool IsCombineSubProcess { get; set; }
        /// <summary>�O�_������AllParts</summary>
        [Required]
        [Display(Name = "")]
        public bool IsNoneShellNoCreateAllParts { get; set; }
        /// <summary>�t��</summary>
        [Required]
        [StringLength(2)]
        [Display(Name = "")]
        public string Region { get; set; }
        /// <summary></summary>
        [Required]
        [Display(Name = "")]
        public int DQSQtyPCT { get; set; }

    }
}
