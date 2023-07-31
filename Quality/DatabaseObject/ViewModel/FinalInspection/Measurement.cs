using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class ServiceMeasurement : BaseResult
    {
        public string FinalInspectionID { get; set; }

        public List<SelectListItem> ListArticle { get; set; }

        public List<ArticleSize> ListSize { get; set; }

        public List<SelectListItem> ListProductType { get; set; }

        public string SelectedArticle { get; set; }

        public string SelectedSize { get; set; }

        public string SelectedProductType { get; set; }

        public string SizeUnit { get; set; }

        public List<MeasurementItem> ListMeasurementItem { get; set; }

    }

    public class MeasurementItem
    {
        public string Description { get; set; }
        public string SizeSpec { get; set; }
        public string Tol1 { get; set; }
        public string Tol2 { get; set; }
        public string Size { get; set; }
        public string ResultSizeSpec { get; set; } = string.Empty;
        public string Code { get; set; }
        public long MeasurementUkey { get; set; }
        public bool CanEdit { get; set; }
    }

    public class ArticleSize
    {
        public string Article { get; set; }
        public string SizeCode { get; set; }
    }

    //public class MeasurementView
    //{
    //    public List<SelectListItem> ListArticle { get; set; }

    //    public List<SelectListItem> ListSize { get; set; }

    //    public List<SelectListItem> ListProductType { get; set; }
    //    public List<MeasurementViewItem> ListMeasurementViewItem { get; set; }
    //}

    public class MeasurementViewItem
    {
        public string Article { get; set; }
        public string Size { get; set; }
        public string ProductType { get; set; }
        public string MeasurementDataByJson { get; set; }
    }
}
