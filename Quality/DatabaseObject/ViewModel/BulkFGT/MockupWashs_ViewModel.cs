using DatabaseObject.ProductionDB;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class MockupWashs_ViewModel : BaseResult
    {
        public List<string> ReportNos { get; set; }

        public List<MockupWash_ViewModel> MockupWash { get; set; }

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

        public List<SelectListItem> TestingMethod_Source { get; set; }
    }

    public class MockupWash_ViewModel : MockupWash
    {
        /// <summary>T1 廠商 名稱</summary>
        [Display(Name = "T1 廠商 名稱")]
        public string T1SubconName { get; set; }

        /// <summary>T2 廠商 名稱</summary>
        [Display(Name = "T2 廠商 名稱")]
        public string T2SupplierName { get; set; }

        /// <summary>技術人員 名稱</summary>
        [Display(Name = "技術人員 名稱")]
        public string TechnicianName { get; set; }

        /// <summary>業務 名稱</summary>
        [Display(Name = "業務 名稱")]
        public string MRName { get; set; }

        /// <summary>LastEditName</summary>
        [Display(Name = "LastEditName")]
        public string LastEditName { get; set; }

        /// <summary>TestingMethodDescription</summary>
        [Display(Name = "TestingMethodDescription")]
        public string TestingMethodDescription { get; set; }

        /// <summary>圖檔完整路徑</summary>
        [Display(Name = "圖檔完整路徑")]
        public string SignaturePic { get; set; }

        [Display(Name = "報表Result")]
        public bool ReportResult { get; set; }

        [Display(Name = "報表Msg")]
        public string ErrorMessage { get; set; }

        [Display(Name = "報表暫存檔名")]
        public string TempFileName { get; set; }

        public List<MockupWash_Detail_ViewModel> MockupWash_Detail { get; set; }
    }

    public class MockupWash_Detail_ViewModel : MockupWash_Detail
    {
        public string ArtworkColorName { get; set; }
        public string FabricColorName { get; set; }
        public string LastUpdate { get; set; }
    }


    public class AccessoryRefNos : BaseResult
    {
        public List<AccessoryRefNo> AccessoryRefNo { get; set; }
    }

    public class AccessoryRefNo
    {
        public string Refno { get; set; }
    }
}
