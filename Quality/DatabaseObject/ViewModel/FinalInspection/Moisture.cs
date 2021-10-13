using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class Moisture : BaseResult
    {
        public string FinalInspectionID { get; set; }
        public decimal? FinalInspection_CTNMoistureStandard { get; set; }
        public List<string> ListArticle { get; set; }
        public List<CartonItem> ListCartonItem { get; set; }
        public List<EndlineMoisture> ListEndlineMoisture { get; set; }
        public List<SelectListItem> ActionSelectListItem { get; set; }
    }

    public class CartonItem
    {
        public long FinalInspection_OrderCartonUkey { get; set; }
        public string Article { get; set; }
        public string OrderID { get; set; }
        public string PackinglistID { get; set; }
        public string CTNNo { get; set; }
    }

    public class MoistureResult
    {
        public string FinalInspectionID { get; set; }
        public string Article { get; set; }
        public long? FinalInspection_OrderCartonUkey { get; set; }
        public string Instrument { get; set; }
        public string Fabrication { get; set; }
        public decimal? GarmentTop { get; set; }
        public decimal? GarmentMiddle { get; set; }
        public decimal? GarmentBottom { get; set; }
        public decimal? CTNInside { get; set; }
        public decimal? CTNOutside { get; set; }
        public string Action { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string AddName { get; set; }
    }

    public class ViewMoistureResult
    {
        public long Ukey { get; set; }
        public string Article { get; set; }
        public string CTNNo { get; set; }
        public string Instrument { get; set; }
        public string Fabrication { get; set; }
        public decimal? GarmentStandard { get; set; }
        public decimal? GarmentTop { get; set; }
        public decimal? GarmentMiddle { get; set; }
        public decimal? GarmentBottom { get; set; }
        public decimal? CTNStandard { get; set; }
        public decimal? CTNInside { get; set; }
        public decimal? CTNOutside { get; set; }
        public string Result { get; set; }
        public string Action { get; set; }
        public string Remark { get; set; }

    }
}
