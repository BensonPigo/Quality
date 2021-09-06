using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class FabricCrkShrkTest_Result
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
        public string CrockingRemark { get; set; }
        public bool? NonHeat { get; set; }
        public string Heat { get; set; }
        public DateTime? HeatDate { get; set; }
        public string HeatRemark { get; set; }
        public bool? NonWash { get; set; }
        public string Wash { get; set; }
        public DateTime? WashDate { get; set; }
        public string WashRemark { get; set; }
        public string ReceivingID { get; set; }
    }

    public class FabricCrkShrkTestCrocking_Result
    {
        public long ID { get; set; }
        public int CrockingTestOption { get; set; }

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
        public string Supp { get; set; }
        public string Crocking { get; set; }
        public DateTime? CrockingDate { get; set; }
        public string StyleID { get; set; }
        public string SCIRefno { get; set; }
        public string Name { get; set; }
        public string BrandID { get; set; }
        public string Refno { get; set; }
        public bool? NonCrocking { get; set; }
        public string DescDetail { get; set; }
        public string CrockingRemark { get; set; }
        public bool CrockingEncdoe { get; set; }
        public byte[] CrockingTestBeforePicture { get; set; }
        public byte[] CrockingTestAfterPicture { get; set; }
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
        public string Remark { get; set; }
        public string LastUpdate { get; set; }
    }

    public class FabricCrkShrkTestHeat_Result
    {
        public long ID { get; set; }
        public FabricCrkShrkTestHeat_Main Heat_Main { get; set; }

        public List<FabricCrkShrkTestHeat_Detail> Heat_Detail { get; set; }
    }

    public class FabricCrkShrkTestHeat_Main
    {
        public string POID { get; set; }
        public string SEQ { get; set; }
        public string ColorID { get; set; }
        public decimal ArriveQty { get; set; }
        public DateTime? WhseArrival { get; set; }
        public string ExportID { get; set; }
        public string Supp { get; set; }
        public string Heat { get; set; }
        public DateTime? HeatDate { get; set; }
        public string StyleID { get; set; }
        public string SCIRefno { get; set; }
        public string Name { get; set; }
        public string BrandID { get; set; }
        public string Refno { get; set; }
        public bool? NonHeat { get; set; }
        public string DescDetail { get; set; }
        public string HeatRemark { get; set; }
        public bool HeatEncode { get; set; }
        public byte[] HeatTestBeforePicture { get; set; }
        public byte[] HeatTestAfterPicture { get; set; }
    }

    public class FabricCrkShrkTestHeat_Detail
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

    public class FabricCrkShrkTestWash_Result
    {
        public long ID { get; set; }
        public FabricCrkShrkTestWash_Main Wash_Main { get; set; }

        public List<FabricCrkShrkTestWash_Detail> Wash_Detail { get; set; }
    }

    public class FabricCrkShrkTestWash_Main
    {
        public string POID { get; set; }
        public string SEQ { get; set; }
        public string ColorID { get; set; }
        public decimal ArriveQty { get; set; }
        public DateTime? WhseArrival { get; set; }
        public string ExportID { get; set; }
        public string Supp { get; set; }
        public string Wash { get; set; }
        public DateTime? WashDate { get; set; }
        public string StyleID { get; set; }
        public string SCIRefno { get; set; }
        public string Name { get; set; }
        public string BrandID { get; set; }
        public string Refno { get; set; }
        public bool? NonWash { get; set; }
        public string SkewnessOptionID { get; set; }
        public string DescDetail { get; set; }
        public string WashRemark { get; set; }
        public byte[] WashTestBeforePicture { get; set; }
        public byte[] WashTestAfterPicture { get; set; }
    }

    public class FabricCrkShrkTestWash_Detail
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
        public decimal SkewnessTest1 { get; set; }
        public decimal SkewnessTest2 { get; set; }
        public decimal SkewnessTest3 { get; set; }
        public decimal SkewnessTest4 { get; set; }
        public decimal SkewnessRate { get; set; }
        public DateTime? Inspdate { get; set; }
        public string Inspector { get; set; }
        public string Name { get; set; }
        public string Remark { get; set; }
        public string LastUpdate { get; set; }
    }
}
