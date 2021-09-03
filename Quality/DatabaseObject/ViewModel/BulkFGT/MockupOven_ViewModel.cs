using DatabaseObject.ProductionDB;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class MockupOven_ViewModel : MockupOven
    {
        public List<string> ReportNo_Source { get; set; }

        public List<SelectListItem> Result_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="Pass",Value="Pass"},
                    new SelectListItem(){ Text="Fail",Value="Fail"},
                };
            }
            set { }
        }

        [Display(Name = "T1 廠商 名稱")]
        public string T1SubconAbb { get; set; }

        [Display(Name = "T2 廠商 名稱")]
        public string T2SupplierAbb { get; set; }

        [Display(Name = "技術人員 名稱")]
        public string TechnicianName { get; set; }

        [Display(Name = "業務 名稱")]
        public string MRName { get; set; }

        [Display(Name = "LastEditName")]
        public string LastEditName { get; set; }

        [Display(Name = "報表電子簽章圖檔完整路徑")]
        public string SignaturePic { get; set; }

        [Display(Name = "報表Result")]
        public bool ReportResult { get; set; }

        [Display(Name = "報表Msg")]
        public string ReportErrorMessage { get; set; }

        [Display(Name = "報表暫存檔名")]
        public string TempFileName { get; set; }

        public List<MockupOven_Detail_ViewModel> MockupOven_Detail { get; set; }
    }

    public class MockupOven_Detail_ViewModel : MockupOven_Detail
    {
        public List<SelectListItem> Result_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="Pass",Value="Pass"},
                    new SelectListItem(){ Text="Fail",Value="Fail"},
                };
            }
            set { }
        }

        public string ArtworkColorName { get; set; }
        public string FabricColorName { get; set; }
        public string LastUpdate { get; set; }
    }
}
