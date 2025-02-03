using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class FabricCrkShrkTest_Result : BaseResult
    {
        public FabricCrkShrkTest_Main Main { get; set; }

        public List<FabricCrkShrkTest_Detail> Details { get; set; }
    }

    public class FabricCrkShrkTest_Main
    {
        public string POID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public DateTime? CutInline { get; set; }
        public DateTime? MinSciDelivery { get; set; }
        public DateTime? TargetLeadTime { get; set; }
        public DateTime? CompletionDate { get; set; }
        public decimal FIRLabInspPercent { get; set; }
        public bool? complete { get; set; }
        public string FirLaboratoryRemark { get; set; }
        public string CreateBy { get; set; }
        public string EditBy { get; set; }
    }

    public class FabricCrkShrkTest_Detail
    {
        public long ID { get; set; }
        public string ReportNo { get; set; }
        public string Seq { get; set; }
        public string WKNo { get; set; }
        public DateTime? WhseArrival { get; set; }
        public string SCIRefno { get; set; }
        public string Refno { get; set; }
        public string ColorID { get; set; }
        public string Supplier { get; set; }
        public decimal ArriveQty { get; set; }
        public DateTime? ReceiveSampleDate { get; set; }
        public DateTime? InspDeadline { get; set; }
        public string AllResult { get; set; }
        public bool? NonCrocking { get; set; }
        public string Crocking { get; set; }
        public DateTime? CrockingDate { get; set; }
        //public string CrockingRemark { get; set; }

        private string _CrockingRemark;

        public string CrockingRemark
        {
            get => _CrockingRemark ?? string.Empty;
            set => _CrockingRemark = value;
        }
        public bool? NonHeat { get; set; }
        public string Heat { get; set; }
        public DateTime? HeatDate { get; set; }
        public string HeatRemark { get; set; }

        public bool? NonIron { get; set; }
        public string Iron { get; set; }
        public DateTime? IronDate { get; set; }
        public string IronRemark { get; set; }

        public bool? NonWash { get; set; }
        public string Wash { get; set; }
        public DateTime? WashDate { get; set; }
        public string WashRemark { get; set; }
        public string ReceivingID { get; set; }
    }

    public class FabricCrkShrkTestCrocking_Result : BaseResult
    {
        public long ID { get; set; }
        public int CrockingTestOption { get; set; }
        public string MDivisionID { get; set; }

        public List<string> ScaleIDs { get; set; }

        public FabricCrkShrkTestCrocking_Main Crocking_Main { get; set; }

        public List<FabricCrkShrkTestCrocking_Detail> Crocking_Detail { get; set; }
    }

    public class FabricCrkShrkTestCrocking_Main
    {
        public string POID { get; set; }
        public string SEQ { get; set; }
        public string ColorID { get; set; }
        public decimal ArriveQty { get; set; }
        public DateTime? WhseArrival { get; set; }
        public string ExportID { get; set; }
        public string ReceivingID { get; set; }
        
        public string Supp { get; set; }
        public string Crocking { get; set; }
        public DateTime? CrockingDate { get; set; }
        public string StyleID { get; set; }
        public string SCIRefno { get; set; }
        public string CrockingInspector { get; set; }
        public string BrandID { get; set; }
        public string Refno { get; set; }
        public bool? NonCrocking { get; set; }
        public string DescDetail { get; set; }

        private string _CrockingRemark;
        public string CrockingRemark
        {
            get => _CrockingRemark ?? string.Empty;
            set => _CrockingRemark = value;
        }
        public bool CrockingEncdoe { get; set; }
        public byte[] CrockingTestPicture1 { get; set; }
        public byte[] CrockingTestPicture2 { get; set; }
        public byte[] CrockingTestPicture3 { get; set; }
        public byte[] CrockingTestPicture4 { get; set; }
        public string MailSubject { get; set; }
        public DateTime? CrockingReceiveDate { get; set; }
        public string CrockingApprover { get; set; } = string.Empty;
        public string CrockingApproverName { get; set; } = string.Empty;
    }

    public class Crocking_Excel
    {
        public string ReportNo { get; set; }
        public string FactoryID { get; set; }
        public DateTime? SubmitDate { get; set; }
        public string SeasonID { get; set; }
        public string BrandID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string POID { get; set; }
        public string Roll { get; set; }
        public string Dyelot { get; set; }
        public string SCIRefno_Color { get; set; }
       public string Refno { get; set; }
        public string Crocking { get; set; }
        public string Color { get; set; }
        public string DryScale { get; set; }
        public string DryScale_Weft { get; set; }
        public string WetScale{ get; set; }
        public string WetScale_Weft { get; set; }
        public string ResultDry { get; set; }
        public string ResultDry_Weft { get; set; }
        public string ResultWet { get; set; }
        public string ResultWet_Weft { get; set; }
        public string Remark { get; set; }
        public string Inspector { get; set; }
        public DateTime? CrockingReceiveDate { get; set; }
        public string CrockingInspector { get; set; }
        public string CrockingInspectorName { get; set; }
        public string CrockingApprover { get; set; }
        public string CrockingApproverName { get; set; }
        public Byte[] InspectorSignature { get; set; }
        public Byte[] ApproverSignature { get; set; }
        public Byte[] CrockingTestPicture1 { get; set; }
        public Byte[] CrockingTestPicture2 { get; set; }
        public Byte[] CrockingTestPicture3 { get; set; }
        public Byte[] CrockingTestPicture4 { get; set; }

    }

    public class FabricCrkShrkTestCrocking_Detail : CompareBase
    {
        public string Roll { get;set; }
        public string Dyelot { get; set; }
        public string Result { get; set; }
        public string DryScale { get; set; }
        public string ResultDry { get; set; }
        public string DryScale_Weft { get; set; }
        public string ResultDry_Weft { get; set; }
        public string WetScale { get; set; }
        public string ResultWet { get; set; }
        public string WetScale_Weft { get; set; }
        public string ResultWet_Weft { get; set; }
        public DateTime? Inspdate { get; set; }
        public string Inspector { get; set; }
        public string Name { get; set; }
        private string _Remark;
        public string Remark
        {
            get => _Remark ?? string.Empty;
            set => _Remark = value;
        }
        public string LastUpdate { get; set; }
    }

    public class FabricCrkShrkTestHeat_Result : BaseResult
    {
        public long ID { get; set; }
        public string MDivisionID { get; set; }
        public FabricCrkShrkTestHeat_Main Heat_Main { get; set; }

        public List<FabricCrkShrkTestHeat_Detail> Heat_Detail { get; set; }
    }

    public class FabricCrkShrkTestIron_Result : BaseResult
    {
        public long ID { get; set; }
        public string MDivisionID { get; set; }
        public FabricCrkShrkTestIron_Main Iron_Main { get; set; }

        public List<FabricCrkShrkTestIron_Detail> Iron_Detail { get; set; }
    }
    public class FabricCrkShrkTestIron_Main
    {
        public string FactoryID { get; set; }
        public string ReportNo { get; set; }
        public string POID { get; set; }
        public string SEQ { get; set; }
        public string ColorID { get; set; }
        public decimal ArriveQty { get; set; }
        public DateTime? WhseArrival { get; set; }
        public string ExportID { get; set; }
        public string ReceivingID { get; set; }
        public string Supp { get; set; }
        public string Iron { get; set; }
        public DateTime? IronDate { get; set; }
        public string StyleID { get; set; }
        public string SCIRefno { get; set; }
        public string IronInspector { get; set; }
        public string IronInspectorName { get; set; }
        public string BrandID { get; set; }
        public string Refno { get; set; }
        public bool? NonIron { get; set; }
        public string DescDetail { get; set; }
        private string _IronRemark;
        public string IronRemark
        {
            get => _IronRemark ?? string.Empty;
            set => _IronRemark = value;
        }
        public DateTime? IronReceiveDate { get; set; }
        public string IronApprover { get; set; } = string.Empty;
        public string IronApproverName { get; set; } = string.Empty;

        public bool IronEncode { get; set; }
        public byte[] IronTestBeforePicture { get; set; }
        public byte[] IronTestAfterPicture { get; set; }
        public byte[] InspectorSignature { get; set; }
        public byte[] ApproverSignature { get; set; }
        public string MailSubject { get; set; }
    }
    public class FabricCrkShrkTestHeat_Main
    {
        public string FactoryID { get; set; }
        public string ReportNo { get; set; }
        public string POID { get; set; }
        public string SEQ { get; set; }
        public string ColorID { get; set; }
        public decimal ArriveQty { get; set; }
        public DateTime? WhseArrival { get; set; }
        public string ExportID { get; set; }
        public string ReceivingID { get; set; }
        public string Supp { get; set; }
        public string Heat { get; set; }
        public DateTime? HeatDate { get; set; }
        public string StyleID { get; set; }
        public string SCIRefno { get; set; }
        public string HeatInspector { get; set; }
        public string HeatInspectorName { get; set; }
        public string BrandID { get; set; }
        public string Refno { get; set; }
        public bool? NonHeat { get; set; }
        public string DescDetail { get; set; }
        private string _HeatRemark;
        public string HeatRemark
        {
            get => _HeatRemark ?? string.Empty;
            set => _HeatRemark = value;
        }
        public bool HeatEncode { get; set; }
        public byte[] HeatTestBeforePicture { get; set; }
        public byte[] HeatTestAfterPicture { get; set; }
        public string MailSubject { get; set; }
        public DateTime? HeatReceiveDate { get; set; }
        //public int Heat_Temperature { get; set; }
        //public int Heat_Second { get; set; }
        //public decimal Heat_Pressure { get; set; }
        public string HeatApprover { get; set; } = string.Empty;
        public string HeatApproverName { get; set; } = string.Empty;
        public byte[] InspectorSignature { get; set; }
        public byte[] ApproverSignature { get; set; }
    }

    public class FabricCrkShrkTestHeat_Detail : CompareBase
    {
        public string Roll { get; set; }
        public string Dyelot { get; set; }
        public decimal HorizontalOriginal { get; set; }
        public decimal VerticalOriginal { get; set; }
        public string Result { get; set; }
        public decimal HorizontalTest1 { get; set; }
        public decimal HorizontalTest2 { get; set; }
        public decimal HorizontalTest3 { get; set; }
        public decimal HorizontalAverage { get; set; }
        public decimal HorizontalRate { get; set; }
        public decimal VerticalTest1 { get; set; }
        public decimal VerticalTest2 { get; set; }
        public decimal VerticalTest3 { get; set; }
        public decimal VerticalAverage { get; set; }
        public decimal VerticalRate { get; set; }
        public DateTime? Inspdate { get; set; }
        public string Inspector { get; set; }
        public string Name { get; set; }
        public string Remark { get; set; }
        public string LastUpdate { get; set; }
    }

    public class FabricCrkShrkTestIron_Detail : CompareBase
    {
        public string Roll { get; set; }
        public string Dyelot { get; set; }
        public decimal HorizontalOriginal { get; set; }
        public decimal VerticalOriginal { get; set; }
        public string Result { get; set; }
        public decimal HorizontalTest1 { get; set; }
        public decimal HorizontalTest2 { get; set; }
        public decimal HorizontalTest3 { get; set; }
        public decimal HorizontalAverage { get; set; }
        public decimal HorizontalRate { get; set; }
        public decimal VerticalTest1 { get; set; }
        public decimal VerticalTest2 { get; set; }
        public decimal VerticalTest3 { get; set; }
        public decimal VerticalAverage { get; set; }
        public decimal VerticalRate { get; set; }
        public DateTime? Inspdate { get; set; }
        public string Inspector { get; set; }
        public string Name { get; set; }
        public string Remark { get; set; }
        public string LastUpdate { get; set; }
    }

    public class FabricCrkShrkTestWash_Result : BaseResult
    {
        public long ID { get; set; }
        public string MDivisionID { get; set; }
        public FabricCrkShrkTestWash_Main Wash_Main { get; set; }

        public List<FabricCrkShrkTestWash_Detail> Wash_Detail { get; set; }
    }

    public class FabricCrkShrkTestWash_Main
    {
        public string FactoryID { get; set; }
        public string ReportNo { get; set; }
        public string POID { get; set; }
        public string SEQ { get; set; }
        public string ColorID { get; set; }
        public decimal ArriveQty { get; set; }
        public DateTime? WhseArrival { get; set; }
        public string ExportID { get; set; }
        public string ReceivingID { get; set; }
        public string Supp { get; set; }
        public string Wash { get; set; }
        public DateTime? WashDate { get; set; }
        public string StyleID { get; set; }
        public string SCIRefno { get; set; }
        public string WashInspector { get; set; }
        public string WashInspectorName { get; set; }
        public string BrandID { get; set; }
        public string Refno { get; set; }
        public bool? NonWash { get; set; }
        public string SkewnessOptionID { get; set; }
        public string DescDetail { get; set; }
        private string _WashRemark;
        public string WashRemark
        {
            get => _WashRemark ?? string.Empty;
            set => _WashRemark = value;
        }
        public bool WashEncode { get; set; }
        public DateTime? WashReceiveDate { get; set; }
        public string WashApprover { get; set; } = string.Empty;
        public string WashApproverName { get; set; } = string.Empty;
        public byte[] WashTestBeforePicture { get; set; }
        public byte[] WashTestAfterPicture { get; set; }

        public byte[] InspectorSignature { get; set; }
        public byte[] ApproverSignature { get; set; }
        public string MailSubject { get; set; }
    }

    public class FabricCrkShrkTestWash_Detail : CompareBase
    {
        public string Roll { get; set; }
        public string Dyelot { get; set; }
        public decimal HorizontalOriginal { get; set; }
        public decimal VerticalOriginal { get; set; }
        public string Result { get; set; }
        public decimal HorizontalTest1 { get; set; }
        public decimal HorizontalTest2 { get; set; }
        public decimal HorizontalTest3 { get; set; }
        public decimal HorizontalAverage { get; set; }
        public decimal HorizontalRate { get; set; }
        public decimal VerticalTest1 { get; set; }
        public decimal VerticalTest2 { get; set; }
        public decimal VerticalTest3 { get; set; }
        public decimal VerticalAverage { get; set; }
        public decimal VerticalRate { get; set; }
        public decimal SkewnessTest1 { get; set; } = 0;
        public decimal SkewnessTest2 { get; set; } = 0;
        public decimal SkewnessTest3 { get; set; } = 0;
        public decimal SkewnessTest4 { get; set; } = 0;
        public decimal SkewnessRate { get; set; }
        public DateTime? Inspdate { get; set; }
        public string Inspector { get; set; }
        public string Name { get; set; }
        public string Remark { get; set; }
        public string LastUpdate { get; set; }
    }
}
