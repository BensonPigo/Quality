using DatabaseObject.ProductionDB;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class MockupCrockings_ViewModel : BaseResult
    {
        public List<string> ReportNos { get; set; }

        public List<MockupCrocking_ViewModel> MockupCrocking { get; set; }

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

        public List<SelectListItem> Scale
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="1",Value="1"},
                    new SelectListItem(){ Text="1-2",Value="1-2"},
                    new SelectListItem(){ Text="2",Value="2"},
                    new SelectListItem(){ Text="2-3",Value="2-3"},
                    new SelectListItem(){ Text="3",Value="3"},
                    new SelectListItem(){ Text="3-4",Value="3-4"},
                    new SelectListItem(){ Text="4",Value="4"},
                    new SelectListItem(){ Text="4-5",Value="4-5"},
                    new SelectListItem(){ Text="5",Value="5"},
                };
            }
            set { }
        }
    }

    public class MockupCrocking_ViewModel : MockupCrocking
    {
        /// <summary>T1 廠商 名稱</summary>
        [Display(Name = "T1 廠商 名稱")]
        public string T1SubconName { get; set; }

        /// <summary>技術人員 名稱</summary>
        [Display(Name = "技術人員 名稱")]
        public string TechnicianName { get; set; }

        /// <summary>業務 名稱</summary>
        [Display(Name = "業務 名稱")]
        public string MRName { get; set; }

        /// <summary>LastEditName</summary>
        [Display(Name = "LastEditName")]
        public string LastEditName { get; set; }

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
        public string ArtworkColorName { get; set; }
        public string FabricColorName { get; set; }
        public string LastUpdate { get; set; }
    }
}
