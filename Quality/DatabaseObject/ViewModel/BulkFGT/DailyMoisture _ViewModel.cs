using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.Public;
using DatabaseObject.RequestModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class DailyMoisture_ViewModel : BaseResult
    {
        public DailyMoisture_Request Request { get; set; }
        public List<SelectListItem> ReportNo_Source { get; set; }
        public List<SelectListItem> DetectionInstrument_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="0",Value="0"},
                    new SelectListItem(){ Text="1",Value="1"},
                    new SelectListItem(){ Text="3",Value="3"},
                    new SelectListItem(){ Text="5",Value="5"},
                    new SelectListItem(){ Text="10",Value="10"},
                    new SelectListItem(){ Text="15",Value="15"},
                    new SelectListItem(){ Text="25",Value="25"},
                };
            }
        }
        public List<EndlineMoisture> EndlineMoisture_Source { get; set; }
        public List<SelectListItem> Action_Source { get; set; }
        public List<SelectListItem> Area_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="Main fabric",Value="Main fabric"},
                    new SelectListItem(){ Text="Pocket seams",Value="Pocket seams"},
                    new SelectListItem(){ Text="Side seam (applicable if garment without pocket)",Value="Side seam (applicable if garment without pocket)"},
                    new SelectListItem(){ Text="Inseam (applicable if garment without pocket)",Value="Inseam (applicable if garment without pocket)"},
                    new SelectListItem(){ Text="Let hem/ bottom hem",Value="Let hem/ bottom hem"},
                    new SelectListItem(){ Text="Neckline/ Collar (back)",Value="Neckline/ Collar (back)"},
                    new SelectListItem(){ Text="Printing (if have)",Value="Printing (if have)"},
                    new SelectListItem(){ Text="Hood seam",Value="Hood seam"},
                    new SelectListItem(){ Text="Elastic (front)",Value="Elastic (front)"},
                    new SelectListItem(){ Text="Elastic (back)",Value="Elastic (back)"},
                    new SelectListItem(){ Text="Elastic (inner)",Value="Elastic (inner)"},
                };
            }
        }
        public List<SelectListItem> Fabric_Source
        {
            get
            {
                return new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="A",Value="A"},
                    new SelectListItem(){ Text="B",Value="B"},
                    new SelectListItem(){ Text="C",Value="C"},
                    new SelectListItem(){ Text="D",Value="D"},
                };
            }
        }
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
        }
        public DailyMoisture_Result Main { get; set; }
        public List<DailyMoisture_Detail_Result> Details { get; set; }
    }

    public class DailyMoisture_Result
    {
        public string ReportNo { get; set; }
        public string OrderID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string Status { get; set; }
        public string Line { get; set; }
        public DateTime? ReportDate { get; set; }
        public string ReportDateText 
        {
            get
            {
                string res = this.ReportDate.HasValue ? this.ReportDate.Value.ToString("yyyy-MM-dd HH:mm:ss") : string.Empty;
                return res;
            }
        }
        public string Instrument { get; set; }
        public string Fabrication { get; set; }
        public decimal Standard { get; set; }
        public string Result { get; set; }
        public string Action { get; set; }
        public string Remark { get; set; }
        public string MRHandleEmail { get; set; }
        public string AddNameText { get; set; }
        public string AddName { get; set; }
        public DateTime? AddDate { get; set; }

        public string EditNameText { get; set; }
        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastEditName
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.EditName}" : string.Empty;
            }
        }
        public string MailSubject { get; set; }

    }
    public class DailyMoisture_Detail_Result : CompareBase
    {
        public string ReportNo { get; set; }
        public long Ukey { get; set; }
        public decimal Point1 { get; set; }
        public decimal Point2 { get; set; }
        public decimal Point3 { get; set; }
        public decimal Point4 { get; set; }
        public decimal Point5 { get; set; }
        public decimal Point6 { get; set; }
        public string Area { get; set; }
        public string Fabric { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string EditNameText { get; set; }
        public string EditName { get; set; }
        public DateTime? EditDate { get; set; }
        public string LastEditName
        {
            get
            {
                return this.EditDate.HasValue ? $"{this.EditDate.Value.ToString("yyyy-MM-dd HH:mm:ss")}-{this.EditName}" : string.Empty;
            }
        }
        public string ReSetResult(decimal Standard)
        {
            this.Result = "Pass";
            if (this.Point1 > Standard || this.Point2 > Standard || this.Point3 > Standard || this.Point4 > Standard || this.Point5 > Standard || this.Point6 > Standard)
            {
                this.Result = "Fail";
            }

            return this.Result;
        }
    }
}
