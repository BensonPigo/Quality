using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.BulkFGT
{
    public class SearchList_ViewModel : ResultModelBase<SearchList_Result>
    {
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string StyleID { get; set; }
        public string Article { get; set; }
        public string Type { get; set; }
        public string TempFileName { get; set; }
        public DateTime? ReceivedDate_s
        {
            get
            {
                return string.IsNullOrEmpty(ReceivedDate_sText) ? (DateTime?)null : DateTime.Parse(ReceivedDate_sText); ;
            }
        }
        public string ReceivedDate_sText { get; set; }
        public DateTime? ReceivedDate_e
        {
            get
            {
                return string.IsNullOrEmpty(ReceivedDate_eText) ? (DateTime?)null : DateTime.Parse(ReceivedDate_eText); ;
            }
        }

        public string ReceivedDate_eText { get; set; }

        public DateTime? ReportDate_s
        {
            get
            {
                return string.IsNullOrEmpty(ReportDate_sText) ? (DateTime?)null : DateTime.Parse(ReportDate_sText); ;
            }
        }
        public string ReportDate_sText { get; set; }
        public DateTime? ReportDate_e
        {
            get
            {
                return string.IsNullOrEmpty(ReportDate_eText) ? (DateTime?)null : DateTime.Parse(ReportDate_eText); ;
            }
        }

        public string ReportDate_eText { get; set; }
        public List<SelectListItem> TypeDatasource { get; set; }

        public string MDivisionID { get; set; }
    }


    public class SearchList_Result
    {
        public string Type { get; set; }
        public string ReportNo { get; set; }
        public string OrderID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string Article { get; set; }
        public string Artwork { get; set; }
        public string Result { get; set; }
        public DateTime? TestDate { get; set; }
        public DateTime? ReceivedDate { get; set; }
        public DateTime? ReleasedDate { get; set; }

    }
}
