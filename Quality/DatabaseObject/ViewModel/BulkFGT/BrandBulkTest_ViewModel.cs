using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class BrandBulkTest_ViewModel : BaseResult
    {
        public BrandBulkTest_Request Request { get; set; }
        public BrandBulkTest Main { get; set; }
        public List<BrandBulkTest> MainList { get; set; }
        public List<BrandBulkTestDox> BrandBulkTestDoxList { get; set; }

        public List<SelectListItem> Article_Source { get; set; }
        public List<SelectListItem> Artwork_Source { get; set; }
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
        public string DownloadFileName { get; set; }
        public string DownloadFileFullName { get; set; }

    }
    public class BrandBulkTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string ReportNo { get; set; }
        public string TestItem { get; set; }
    }

    public class BrandBulkTest
    {
        public string ReportNo { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string FactoryID { get; set; }
        public string OrderID { get; set; }
        public DateTime? ReportDate { get; set; }
        public string ReportDateText
        {
            get
            {
                string rtn = string.Empty;
                if (this.ReportDate.HasValue)
                {
                    rtn = this.ReportDate.Value.ToString("yyyy/MM/dd");
                }
                return rtn;
            }

        }
        public long TestItemUkey { get; set; }
        public string TestItem { get; set; }
        public string FabricRefno { get; set; }
        public string FabricColor { get; set; }
        public string AccessoryRefno { get; set; }
        public string AccessoryColor { get; set; }
        public string Artwork { get; set; }
        public string ArtworkRefno { get; set; }
        public string ArtworkColor { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string AddName { get; set; }
        public DateTime? AddDate { get; set; }

        public string CreateBy
        {
            get
            {
                return this.AddDate.HasValue ? $"{this.AddDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.AddName}" : string.Empty;
            }
        }

        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastUpadate
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.EditName}" : string.Empty;
            }
        }

        public string EditType { get; set; }
    }

    public class BrandBulkTestItem
    {
        public long Ukey { get; set; }
        public string BrandID { get; set; }
        public string Category { get; set; }
        public string TestClassify { get; set; }
        public string DocType { get; set; }
        public string TestItem { get; set; }
    }
    public class BrandBulkTestDox : CompareBase
    {
        public string ReportNo { get; set; }
        public long Ukey { get; set; }

        /// <summary>
        /// DB帶出的資料都會是true，User畫面上傳新增的都會是false，以此來判斷是否需要上傳該筆檔案
        /// </summary>
        public bool IsOldFile{ get; set; }
        public string FileName { get; set; }
        public HttpPostedFileBase BrandBulkTestDoxFile { get; set; }

        public string FilePath { get; set; }
        public string AddName { get; set; }
        public DateTime? AddDate { get; set; }

        public string CreateBy
        {
            get
            {
                return this.AddDate.HasValue ? $"{this.AddDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.AddName}" : string.Empty;
            }
        }

        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastUpadate
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.EditName}" : string.Empty;
            }
        }

        public string EditType { get; set; }
    }
}
