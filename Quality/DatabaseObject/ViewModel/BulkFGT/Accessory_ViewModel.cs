using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class Accessory_ViewModel : ResultModelBase<Accessory_Result>
    {
        public string ReqOrderID { get; set; }

        //表頭
        public string OrderID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public DateTime? EarliestCutDate { get; set; }
        public DateTime? EarliestSCIDel { get; set; }
        public DateTime? TargetLeadTime { get; set; }
        public DateTime? CompletionDate { get; set; }
        public decimal ArticlePercent { get; set; }
        public bool MtlCmplt { get; set; }
        public string Remark { get; set; }
        public string CreateBy { get; set; }
        public string EditBy { get; set; }

    }

    public class Accessory_Result
    {
        //表身
        public Int64 AIR_LaboratoryID { get; set; }
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Seq { get; set; }

        public string ExportID { get; set; }
        public DateTime? WhseArrival { get; set; }
        public string SCIRefno { get; set; }
        public string Refno { get; set; }
        public string SuppID { get; set; }
        public string Supplier { get; set; }
        public string ColorID { get; set; }
        public string SizeSpec { get; set; }
        public decimal ArriveQty { get; set; }
        public DateTime? InspDeadline { get; set; }
        public string OverAllResult { get; set; } // Pass / Fail

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
        public string ReceivingID { get; set; }
        //

        public bool NonWashingFastness { get; set; }
        public string WashingFastnessResult { get; set; }
        public string WashingFastnessScale { get; set; }
        public string WashingFastnessInspector { get; set; }
        public string WashingFastnessRemark { get; set; }
    }


    public class Accessory_Oven
    {
        //Oven

        public Int64 AIR_LaboratoryID { get; set; }
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Seq { get; set; }
        public string POID { get; set; }
        public string OverAllResult { get; set; } // Pass / Fail

        public string SCIRefno { get; set; }
        public string WKNo { get; set; }
        public string Refno { get; set; }
        public decimal ArriveQty { get; set; }
        public string Supplier { get; set; }
        public string Unit { get; set; }
        public string Color { get; set; }
        public string Size { get; set; }
        public string Scale { get; set; }
        public string OvenResult { get; set; }
        public string Remark { get; set; }
        public string OvenInspector { get; set; }
        public string OvenInspectorName { get; set; }
        public DateTime? OvenDate { get; set; }
        public string EditName { get; set; }

        public bool OvenEncode { get; set; }


        public Byte[] OvenTestBeforePicture { get; set; }
        public Byte[] OvenTestAfterPicture { get; set; }

        public List<SelectListItem> ScaleData { get; set; }
        public List<SelectListItem> ResultData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="",Value=""},
                    new SelectListItem(){ Text="Pass",Value="Pass"},
                    new SelectListItem(){ Text="Fail",Value="Fail"},
                };
            }
            set { }
        }

        public bool Result { get; set; } = true;
        public string ErrorMessage { get; set; }


        // Result = Fail時，要收信的人
        public string ToAddress { get; set; }
        public string CcAddress { get; set; }
    }

    public class Accessory_OvenExcel : Accessory_Oven
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
    }

    public class Accessory_Wash
    {
        public Int64 AIR_LaboratoryID { get; set; }
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Seq { get; set; }
        public string POID { get; set; }
        public string OverAllResult { get; set; } // Pass / Fail

        public string SCIRefno { get; set; }
        public string WKNo { get; set; }
        public string Refno { get; set; }

        public decimal ArriveQty { get; set; }
        public string SuppID { get; set; }
        public string Supplier { get; set; }
        public string Unit { get; set; }
        public string Color { get; set; }
        public string Size { get; set; }
        public string Scale { get; set; }
        public string WashResult { get; set; }
        public string Remark { get; set; }
        public string WashInspector { get; set; }
        public string WashInspectorName { get; set; }

        public string MachineWash { get; set; }
        public int WashingTemperature { get; set; }
        public string DryProcess { get; set; }
        public string MachineModel { get; set; }
        public int WashingCycle { get; set; }

        public DateTime? WashDate { get; set; }
        public string EditName { get; set; }

        public bool WashEncode { get; set; }

        public bool Result { get; set; } = true;
        public string ErrorMessage { get; set; }

        public Byte[] WashTestBeforePicture { get; set; }
        public Byte[] WashTestAfterPicture { get; set; }

        public List<SelectListItem> ScaleData { get; set; }
        public List<SelectListItem> ResultData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="",Value=""},
                    new SelectListItem(){ Text="Pass",Value="Pass"},
                    new SelectListItem(){ Text="Fail",Value="Fail"},
                };
            }
            set { }
        }
        public List<SelectListItem> MachineWashData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="" ,Value = ""},
                    new SelectListItem(){ Text="Top load",Value="Top load"},
                    new SelectListItem(){ Text="Front load",Value="Front load"},
                };
            }
            set { }
        }
        public List<SelectListItem> WashingTemperatureData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="0" ,Value="0"},
                    new SelectListItem(){ Text="30" ,Value="30"},
                    new SelectListItem(){ Text="40" ,Value="40"},
                    new SelectListItem(){ Text="60" ,Value="60"},
                };
            }
            set { }
        }
        public List<SelectListItem> DryProcessData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="" ,Value=""},
                    new SelectListItem(){ Text="Tumble dry low" ,Value="Tumble dry low"},
                    new SelectListItem(){ Text="Tumble dry Medium" ,Value="Tumble dry Medium"},
                    new SelectListItem(){ Text="Line dry" ,Value="Line dry"},
                };
            }
            set { }
        }
        public List<SelectListItem> MachineModelData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="" ,Value=""},
                    new SelectListItem(){ Text="Miele" ,Value="Miele"},
                    new SelectListItem(){ Text="Whirlpool" ,Value="Whirlpool"},
                };
            }
            set { }
        }
        public List<SelectListItem> WashingCycleData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="0" ,Value="0"},
                    new SelectListItem(){ Text="3" ,Value="3"},
                    new SelectListItem(){ Text="5" ,Value="5"},
                    new SelectListItem(){ Text="10" ,Value="10"},
                    new SelectListItem(){ Text="15" ,Value="15"},
                    new SelectListItem(){ Text="20" ,Value="20"},
                    new SelectListItem(){ Text="25" ,Value="25"},
                };
            }
            set { }
        }

        // Result = Fail時，要收信的人
        public string ToAddress { get; set; }
        public string CcAddress { get; set; }
    }

    public class Accessory_WashExcel : Accessory_Wash
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
    }

    public class Accessory_WashingFastness
    {
        public Int64 AIR_LaboratoryID { get; set; }
        public string Seq1 { get; set; }
        public string Seq2 { get; set; }
        public string Seq { get; set; }
        public string POID { get; set; }
        public string OverAllResult { get; set; } // Pass / Fail

        public string SCIRefno { get; set; }
        public string WKNo { get; set; }
        public string Refno { get; set; }

        public decimal ArriveQty { get; set; }
        public string SuppID { get; set; }
        public string Supplier { get; set; }
        public string Unit { get; set; }
        public string Size { get; set; }
        public string Color { get; set; }

        public bool NonWashingFastness { get; set; }
        public string WashingFastnessResult { get; set; }
        public bool WashingFastnessEncode { get; set; }
        public string WashingFastnessScale { get; set; }
        public string WashingFastnessInspector { get; set; }
        public string WashingFastnessInspectorName { get; set; }
        public DateTime? WashingFastnessReceivedDate { get; set; }
        public DateTime? WashingFastnessReportDate { get; set; }
        public string WashingFastnessRemark { get; set; }

        public string ChangeScale { get; set; }
        public string ResultChange { get; set; }
        public string AcetateScale { get; set; }
        public string ResultAcetate { get; set; }
        public string CottonScale { get; set; }
        public string ResultCotton { get; set; }
        public string NylonScale { get; set; }
        public string ResultNylon { get; set; }
        public string PolyesterScale { get; set; }
        public string ResultPolyester { get; set; }
        public string AcrylicScale { get; set; }
        public string ResultAcrylic { get; set; }
        public string WoolScale { get; set; }
        public string ResultWool { get; set; }
        public string CrossStainingScale { get; set; }
        public string ResultCrossStaining { get; set; }
        public string EditName { get; set; }

        //public bool WashEncode { get; set; }

        public bool Result { get; set; } = true;
        public string ErrorMessage { get; set; }

        public Byte[] WashingFastnessTestBeforePicture { get; set; }
        public Byte[] WashingFastnessTestAfterPicture { get; set; }

        public List<SelectListItem> ScaleData { get; set; }
        public List<SelectListItem> ResultData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="",Value=""},
                    new SelectListItem(){ Text="Pass",Value="Pass"},
                    new SelectListItem(){ Text="Fail",Value="Fail"},
                };
            }
            set { }
        }
        public List<SelectListItem> MachineWashData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="" ,Value = ""},
                    new SelectListItem(){ Text="Top load",Value="Top load"},
                    new SelectListItem(){ Text="Front load",Value="Front load"},
                };
            }
            set { }
        }
        public List<SelectListItem> WashingTemperatureData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="0" ,Value="0"},
                    new SelectListItem(){ Text="30" ,Value="30"},
                    new SelectListItem(){ Text="40" ,Value="40"},
                    new SelectListItem(){ Text="60" ,Value="60"},
                };
            }
            set { }
        }
        public List<SelectListItem> DryProcessData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="" ,Value=""},
                    new SelectListItem(){ Text="Tumble dry low" ,Value="Tumble dry low"},
                    new SelectListItem(){ Text="Tumble dry Medium" ,Value="Tumble dry Medium"},
                    new SelectListItem(){ Text="Line dry" ,Value="Line dry"},
                };
            }
            set { }
        }
        public List<SelectListItem> MachineModelData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="" ,Value=""},
                    new SelectListItem(){ Text="Miele" ,Value="Miele"},
                    new SelectListItem(){ Text="Whirlpool" ,Value="Whirlpool"},
                };
            }
            set { }
        }
        public List<SelectListItem> WashingCycleData
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="0" ,Value="0"},
                    new SelectListItem(){ Text="3" ,Value="3"},
                    new SelectListItem(){ Text="5" ,Value="5"},
                    new SelectListItem(){ Text="10" ,Value="10"},
                    new SelectListItem(){ Text="15" ,Value="15"},
                    new SelectListItem(){ Text="20" ,Value="20"},
                    new SelectListItem(){ Text="25" ,Value="25"},
                };
            }
            set { }
        }

        // Result = Fail時，要收信的人
        public string ToAddress { get; set; }
        public string CcAddress { get; set; }
    }

    public class Accessory_WashingFastnessExcel : Accessory_WashingFastness
    {
        public string FactoryID { get; set; }
        public string Article { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string Conclusions { get; set; }
        public string StyleID { get; set; }

        public string Prepared { get; set; }
        public string Executive { get; set; }
    }
}
