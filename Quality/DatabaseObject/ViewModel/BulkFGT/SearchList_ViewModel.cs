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

        public List<SelectListItem> TypeDatasource { get; set; }
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

    }
}
