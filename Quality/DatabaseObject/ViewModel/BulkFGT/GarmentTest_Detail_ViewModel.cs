using DatabaseObject.ProductionDB;
using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class GarmentTest_Detail_ViewModel : CompareBase
    {
        public string DrySelect { get; set; }
        public string Above50 { get; set; }
        public string ReportNo { get; set; }
        public string NeckSelect { get; set; }
        public string GarmentTest_Detail_Inspector { get; set; }
        public string GarmentTest_Detail_AddName { get; set; }
        public string GarmentTest_Detail_EditName { get; set; }
        public Int64? ID { get; set; }
        public int? No { get; set; }
        public string Remark { get; set; }
        public DateTime? SubmitDate { get; set; }
        public string SizeCode { get; set; }
        public string MtlTypeID { get; set; }
        public string OrderID { get; set; }
        public bool NonSeamBreakageTest { get; set; }
        public string Result { get; set; }
        public DateTime? inspdate { get; set; }
        public string inspector { get; set; }
        public string Sender { get; set; }
        public DateTime? SendDate { get; set; }
        public string Receiver { get; set; }
        public DateTime? ReceiveDate { get; set; }
        public string AddName { get; set; }
        public DateTime? AddDate { get; set; }
        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string OldUkey { get; set; }
        public int? ArrivedQty { get; set; }
        public bool? LineDry { get; set; }
        public int? Temperature { get; set; }
        public bool? TumbleDry { get; set; }
        public string Machine { get; set; }
        public bool? HandWash { get; set; }
        public string Composition { get; set; }
        public bool? Neck { get; set; }
        public string Status { get; set; }
        public string LOtoFactory { get; set; } = "";
        public string FabricationType { get; set; }        
        public bool? Above50NaturalFibres { get; set; }
        public bool? Above50SyntheticFibres { get; set; }
        public string SeamBreakageResult { get; set; }
        public string OdourResult { get; set; }
        public string WashResult { get; set; }
        public Byte[] TestBeforePicture { get; set; }
        public Byte[] TestAfterPicture { get; set; }
    }
}
