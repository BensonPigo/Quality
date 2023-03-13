using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class AgingHydrolysisTest_ViewModel : BaseResult
    {
        public AgingHydrolysisTest_Request Request { get; set; }
        public AgingHydrolysisTest_Main MainData { get; set; }
        public List<AgingHydrolysisTest_Main> MainList { get; set; }
        public List<AgingHydrolysisTest_Detail> DetailList { get; set; }

        public List<SelectListItem> OrderID_Source { get; set; }
        public List<SelectListItem> Article_Source { get; set; }

        public List<SelectListItem> MaterialType_Source = new List<SelectListItem>()
        {
            new SelectListItem(){ Text = "Garment" ,Value = "Garment" },
            new SelectListItem(){ Text = "Mockup" ,Value = "Mockup" },
        };
    }
    public class AgingHydrolysisTest_Request
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public long AgingHydrolysisTestID { get; set; }
        public string ReportNo { get; set; }
    }
    public class AgingHydrolysisTest_Main
    {
        public long ID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string FactoryID { get; set; }
        public string MDivisionID { get; set; }
        public string OrderID { get; set; }

        public decimal Tempereture { get; set; }
        public decimal Time { get; set; }
        public decimal Humidity { get; set; }

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

    public class AgingHydrolysisTest_Detail : CompareBase
    {
        public long AgingHydrolysisTestID { get; set; }
        public string ReportNo { get; set; }
        public string MaterialType { get; set; }
        public DateTime? BuyerDelivery { get; set; }
        public string BuyerDeliveryText
        {
            get
            {
                string rtn = string.Empty;
                if (this.BuyerDelivery.HasValue)
                {
                    rtn = this.BuyerDelivery.Value.ToString("yyyy/MM/dd");
                }
                return rtn;
            }

        }
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
        public DateTime? ReceivedDate { get; set; }
        public string ReceivedDateText
        {
            get
            {
                string rtn = string.Empty;
                if (this.ReceivedDate.HasValue)
                {
                    rtn = this.ReceivedDate.Value.ToString("yyyy/MM/dd");
                }
                return rtn;
            }

        }
        public string FabricRefNo { get; set; }
        public string AccRefNo { get; set; }
        public string FabricColor { get; set; }
        public string AccColor { get; set; }
        public string Status { get; set; }
        public string Result { get; set; }
        public string Comment { get; set; }

        public byte[] TestBeforePicture { get; set; }
        public byte[] TestAfterPicture { get; set; }

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


        // From AgingHydrolysisTest
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public decimal Tempereture { get; set; }
        public decimal Time { get; set; }
        public decimal Humidity { get; set; }
    }
    public class AgingHydrolysisTest_Detail_ViewModel : BaseResult
    {
        public AgingHydrolysisTest_Detail MainDetailData { get; set; }
        public List<AgingHydrolysisTest_Detail> MainDetailDataList { get; set; }
        public List<AgingHydrolysisTest_Detail_Mockup> MockupList { get; set; }
        public List<SelectListItem> Scale_Source { get; set; }
        public List<SelectListItem> MaterialType_Source = new List<SelectListItem>()
        {
            new SelectListItem(){ Text = "Garment" ,Value = "Garment" },
            new SelectListItem(){ Text = "Mockup" ,Value = "Mockup" },
        };
        public List<SelectListItem> Result_Source = new List<SelectListItem>()
        {
            new SelectListItem(){Text="Pass",Value="Pass" },
            new SelectListItem(){Text="Fail",Value="Fail" },
        };
    }
    public class AgingHydrolysisTest_Detail_Mockup : CompareBase
    {
        public string ReportNo { get; set; }
        public long Ukey { get; set; }
        public string SpecimenName { get; set; }
        public string ChangeScale { get; set; }
        public string ChangeResult { get; set; }
        public string StainingScale { get; set; }
        public string StainingResult { get; set; }
        public string Comment { get; set; }
        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastUpadate
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.EditName}" : string.Empty;
            }
        }

    }
}
