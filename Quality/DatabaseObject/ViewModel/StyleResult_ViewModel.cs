using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class StyleResult_ViewModel
    {
        /// <summary>
        /// 以Barcode顯示
        /// </summary>
        public string StyleUkey { get; set; }

        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string ProgramID { get; set; }
        public string Description { get; set; }
        public string StyleName { get; set; }
        public string SpecialMark { get; set; }
        public string SMR { get; set; }
        public string Handle { get; set; }
        public string RFT { get; set; }
        public string ProductType { get; set; }
        public string Article { get; set; }

        public string MsgScript { get; set; }

        public List<StyleResult_SampleRFT> SampleRFT { get; set; }
        public List<StyleResult_SampleTesting> SampleTesting { get; set; }
        public List<StyleResult_Exception> Exception { get; set; }
    }

    public class StyleResult_SampleRFT
    {
        public string SP { get; set; }
        public string SampleStage { get; set; }
        public string Factory { get; set; }
        public DateTime? Delivery { get; set; }
        public DateTime? SCIDelivery { get; set; }
        public int InspectedQty { get; set; }
        public decimal RFT { get; set; }
        public int BAProduct { get; set; }
        public decimal BAAuditCriteria { get; set; }
    }

    public class StyleResult_SampleTesting
    {
        public string Article { get; set; }

        public bool? GarmentWash { get; set; }
        public bool? MockupCrockingTest { get; set; }
        public bool? MockupOvenTest { get; set; }
        public bool? MockupWashTest { get; set; }
        public bool? FabricOvenTest { get; set; }

        public bool? Artwork3rdTest_A01 { get; set; }
        public bool? Artwork3rdTest_Physical { get; set; }
    }

    public class StyleResult_Exception
    {
        // 目前看不懂，等Jimmy後續說明
    }
}
