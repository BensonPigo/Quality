using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class Accessory_ViewModel : ResultModelBase<Accessory_Result>
    {
        //表頭
        public string OrderID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public DateTime? EarliestDate { get; set; }
        public DateTime? EarliestSCIDel { get; set; }
        public DateTime? TargetLeadTime { get; set; }
        public DateTime? CompletionDate { get; set; }
        public decimal ArticlePercent { get; set; }
        public string MtlCmplt { get; set; }
        public string Remark { get; set; }
        public string CreateBy { get; set; }
        public string EditBy { get; set; }

    }

    public class Accessory_Result
    {
        //表身
        public string AIR_LaboratoryID { get; set; }
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Seq { get; set; }

        public string WKNo { get; set; }
        public DateTime? WhseArrival { get; set; }
        public string SCIRefno { get; set; }
        public string Refno { get; set; }
        public string SuppID { get; set; }
        public string Supplier { get; set; }
        public string Color { get; set; }
        public string Size { get; set; }
        public int ArriveQty { get; set; }
        public DateTime? InspDeadline { get; set; }
        public int Result { get; set; }

        public bool NonOven { get; set; }
        public string OvenResult { get; set; }
        public string OvenScale { get; set; }
        public DateTime? OvenDate { get; set; }
        public string OvenInspector { get; set; }
        public string OvenRemark { get; set; }

        public bool NonWash { get; set; }
        public string WashResult { get; set; }
        public string WashScale { get; set; }
        public DateTime? WashDate { get; set; }
        public string WashInspector { get; set; }
        public string WashRemark { get; set; }
        public string Receiving { get; set; }
    }


    public class Accessory_Oven
    {
        //Oven

        public string AIR_LaboratoryID { get; set; }
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Seq { get; set; }
        public string POID { get; set; }

        public string SCIRefno { get; set; }
        public string WKNo { get; set; }
        public string Refno { get; set; }
        public int ArriveQty { get; set; }
        public string SuppID { get; set; }
        public string Supplier { get; set; }
        public string Unit { get; set; }
        public string Color { get; set; }
        public string Size { get; set; }
        public string Scale { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string OvenInspector { get; set; }
        public DateTime? OvenDate { get; set; }

        public bool OvenEncode { get; set; }
    }

    public class Accessory_Wash
    {
        public string AIR_LaboratoryID { get; set; }
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Seq { get; set; }
        public string POID { get; set; }
        public string SCIRefno { get; set; }
        public string WKNo { get; set; }
        public string Refno { get; set; }
        public int ArriveQty { get; set; }
        public string SuppID { get; set; }
        public string Supplier { get; set; }
        public string Unit { get; set; }
        public string Color { get; set; }
        public string Size { get; set; }
        public string Scale { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string WashInspector { get; set; }
        public DateTime? WashDate { get; set; }

        public bool WashEncode { get; set; }
    }
}
