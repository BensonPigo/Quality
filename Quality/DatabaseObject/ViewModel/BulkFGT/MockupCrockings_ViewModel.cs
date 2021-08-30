using DatabaseObject.ProductionDB;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace DatabaseObject.ViewModel
{
    public class MockupCrockings_ViewModel : BaseResult
    {
        public List<string> ReportNo { get; set; }

        public string TempFileName { get; set; }

        public List<MockupCrocking_ViewModel> MockupCrocking { get; set; }
    }


    public class MockupCrocking_ViewModel : MockupCrocking
    {
        /// <summary>T1 廠商 名稱</summary>
        [Display(Name = "T1 廠商 名稱")]
        public string Abb { get; set; }
        /// <summary>技術人員 名稱</summary>
        [Display(Name = "技術人員 名稱")]
        public string TechnicianName { get; set; }
        /// <summary>圖檔完整路徑</summary>
        [Display(Name = "圖檔完整路徑")]
        public string SignaturePic { get; set; }

        [Display(Name = "報表Result")]
        public bool ReportResult { get; set; }

        [Display(Name = "報表Msg")]
        public string ErrorMessage { get; set; }

        [Display(Name = "報表暫存檔名")]
        public string TempFileName { get; set; }

        public List<MockupCrocking_Detail_ViewModel> MockupCrocking_Detail { get; set; }
    }

    public class MockupCrocking_Detail_ViewModel: MockupCrocking_Detail
    {
        /// <summary>FabricColorName</summary>
        [Display(Name = "FabricColorName")]
        public string FabricColorName { get; set; }
    }
}
