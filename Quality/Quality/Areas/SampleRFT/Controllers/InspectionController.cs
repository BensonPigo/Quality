using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class InspectionController : BaseController
    {
        private IInspectionService _InspectionService;
        public InspectionController()
        {
            _InspectionService = new InspectionService();
            if (this.SelectItemData == null || this.SelectItemData.Count() == 0)
                this.SelectItemData = _InspectionService.GetSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID }).ToList();

            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.Inspection,,";
        }

        // GET: SampleRFT/Inspection
        public ActionResult Index()
        {
            this.WorkDate = DateTime.Now;
            Inspection_ViewModel setting = new Inspection_ViewModel()
            {
                FactoryID = this.FactoryID,
                Line = this.Line,
                InspectionDate = this.WorkDate,
            };

            List<SelectListItem> FactoryitemList = new SetListItem().ItemListBinding(this.Factorys);
            List<SelectListItem> LineList = new SetListItem().ItemListBinding(this.Lines);
            List<SelectListItem> BrandList = new SetListItem().ItemListBinding(this.Factorys);

            ViewBag.FactoryList = FactoryitemList;
            ViewBag.LineList = LineList;
            ViewBag.BrandList = BrandList;
            ViewBag.UserID = this.UserID;
            ViewBag.ShowSwitch = 1;
            return View(setting);
        }

        [HttpPost]
        public ActionResult Index(Inspection_ViewModel request)
        {
            this.FactoryID = request.FactoryID;
            this.Line = request.Line;            
            this.WorkDate = request.InspectionDate = CheckWorkDate.Check(request.InspectionDate);
            this.SelectItemData = _InspectionService.GetSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID }).ToList();
            List<SelectListItem> FactoryitemList = new SetListItem().ItemListBinding(this.Factorys);
            List<SelectListItem> LineList = new SetListItem().ItemListBinding(this.Lines);
            List<SelectListItem> BrandList = new SetListItem().ItemListBinding(this.Factorys);

            ViewBag.FactoryList = FactoryitemList;
            ViewBag.LineList = LineList;
            ViewBag.BrandList = BrandList;
            ViewBag.UserID = this.UserID;
            ViewBag.ShowSwitch = 0;
            return View(request);
        }

        [HttpPost]
        public ActionResult GetStyle(string OrderID)
        {            
            List<string> styles = this.SelectItemData.Where(x => string.IsNullOrEmpty(OrderID) || x.OrderID.Equals(OrderID))
                                    .GroupBy(x => x.StyleID)
                                    .Select(x => x.Key).ToList();
            string html = "";
            foreach (string style in styles)
            {
                html += "<li><a href='#'>" + style + "</a></li>";
            }

            return Content(html);
        }

        [HttpPost]
        public JsonResult CheckStyle(string StyleID)
        {
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(StyleID))
            {
                return Json(viewModel);
            }

            var query = _InspectionService
                            .GetSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, StyleID = StyleID })
                            .ToList();

            viewModel.Result = query.Any();
            viewModel.StyleID = query.Any() ? query.FirstOrDefault().StyleID : string.Empty;
            return Json(viewModel);
        }

       [HttpPost]
        public ActionResult GetSP(string StyleID)
        {
            List<string> orderIDs = SelectItemData.Where(x => string.IsNullOrEmpty(StyleID) || x.StyleID.Equals(StyleID))
                                    .GroupBy(x => x.OrderID)
                                    .Select(x => x.Key).ToList();

            string html = "";
            foreach (string orderID in orderIDs)
            {
                html += "<li><a href='#'>" + orderID + "</a></li>";
            }

            return Content(html);
        }

        [HttpPost]
        public JsonResult CheckSP(string StyleID, string OrderID)
        {
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(OrderID))
            {
                return Json(viewModel);
            }

            List<Inspection_ViewModel> result = _InspectionService.CheckSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, StyleID = StyleID, OrderID = OrderID }
                                                                                    , InspectionService.SelectType.OrderID)
                                                                   .ToList();
            viewModel = result.Any() ? result.FirstOrDefault() : viewModel;
            viewModel.Result = result.Any();
            return Json(viewModel);
        }

        [HttpPost]
        public ActionResult GetArticle(string StyleID, string OrderID)
        {
            List<string> articles = _InspectionService
                                    .GetSelectItemData(new Inspection_ViewModel() { StyleID = StyleID, OrderID = OrderID })
                                    .GroupBy(x => x.Article)
                                    .Select(x => x.Key).ToList();
            string html = "";
            foreach (string article in articles)
            {
                html += "<li><a href='#'>" + article + "</a></li>";
            }

            return Content(html);
        }

        [HttpPost]
        public JsonResult CheckArticle(string StyleID, string OrderID, string Article)
        {
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(Article))
            {
                return Json(viewModel);
            }

            List<Inspection_ViewModel> result = _InspectionService.CheckSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, StyleID = StyleID, OrderID = OrderID, Article = Article }
                                                                                    , InspectionService.SelectType.Article)
                                                                   .ToList();
            viewModel = result.Any() ? result.FirstOrDefault() : viewModel;
            viewModel.Result = result.Any();
            return Json(viewModel);
        }

        [HttpPost]
        public ActionResult GetSize(string StyleID, string OrderID, string Article)
        {
            List<string> sizes = _InspectionService
                                    .GetSelectItemData(new Inspection_ViewModel() { StyleID = StyleID, OrderID = OrderID, Article = Article })
                                    .GroupBy(x => x.Size)
                                    .Select(x => x.Key).ToList();
            string html = "";
            foreach (string size in sizes)
            {
                html += "<li><a href='#'>" + size + "</a></li>";
            }

            return Content(html);
        }

        [HttpPost]
        public JsonResult CheckSize(string StyleID, string OrderID, string Article, string Size)
        {
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(Size))
            {
                return Json(viewModel);
            }

            List<Inspection_ViewModel> result = _InspectionService.CheckSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, StyleID = StyleID, OrderID = OrderID, Article = Article, Size = Size }
                                                                                    , InspectionService.SelectType.Size)
                                                                   .ToList();
            viewModel = result.Any() ? result.FirstOrDefault() : viewModel;
            viewModel.Result = result.Any();
            return Json(viewModel);
        }

        [HttpPost]
        public ActionResult GetProductType(string StyleID, string OrderID, string Article, string Size)
        {
            List<string> productTypes = _InspectionService
                                    .GetSelectItemData(new Inspection_ViewModel() { StyleID = StyleID, OrderID = OrderID, Article = Article, Size = Size })
                                    .GroupBy(x => x.ProductType)
                                    .Select(x => x.Key).ToList();
            string html = "";
            foreach (string productType in productTypes)
            {
                html += "<li><a href='#'>" + productType + "</a></li>";
            }

            return Content(html);
        }

        [HttpPost]
        public JsonResult CheckProductType(string StyleID, string OrderID, string Article, string Size, string ProductType)
        {
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(ProductType))
            {
                return Json(viewModel);
            }

            List<Inspection_ViewModel> result = _InspectionService.CheckSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, StyleID = StyleID, OrderID = OrderID, Article = Article, Size = Size, ProductType = ProductType }
                                                                                    , InspectionService.SelectType.ProductType)
                                                                   .ToList();
            viewModel = result.Any() ? result.FirstOrDefault() : viewModel;
            viewModel.Result = result.Any();
            return Json(viewModel);
        }

        [HttpPost]
        public JsonResult RefreshQty(string StyleID, string OrderID, string Article, string Size, string ProductType)
        {
            Inspection_ViewModel result = _InspectionService
                    .GetTop3(new Inspection_ViewModel() { FactoryID = this.FactoryID, Line = this.Line, InspectionDate = this.WorkDate });

            if (string.IsNullOrEmpty(StyleID) ||
                string.IsNullOrEmpty(OrderID) ||
                string.IsNullOrEmpty(Article) ||
                string.IsNullOrEmpty(Size) ||
                string.IsNullOrEmpty(ProductType))
            {
                result.Result = false;
            }
            else
            {
                //  get Size Qty Order Qty 
                List<Inspection_ViewModel> inspection_Views = _InspectionService.CheckSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, StyleID = StyleID, OrderID = OrderID, Article = Article, Size = Size, ProductType = ProductType }
                                                                        , InspectionService.SelectType.ProductType)
                                                                .ToList();
                result.Result = inspection_Views.Any();
                Inspection_ViewModel inspection = inspection_Views.FirstOrDefault();
                result.SizeQty = inspection.SizeQty;
                result.SizeBalanceQty = inspection.SizeBalanceQty;
                result.OrderQty = inspection.OrderQty;
                result.OrderBalanceQty = inspection.OrderBalanceQty;
            }

            return Json(result);
        }

        [HttpPost]
        public JsonResult Pass(string StyleID, string OrderID, string Article, string Size, string ProductType)
        {
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            viewModel.Result = true; 
            return Json(viewModel);
        }

        [HttpPost]
        public ActionResult GetDefectType()
        {
            List<GarmentDefectType> garmentDefectTypes = _InspectionService.GetGarmentDefectType().ToList();
            //List<GarmentDefectType> result = garmentDefectTypes.OrderBy(x => x.Seq)
            //            .Select(x => new GarmentDefectType  { ID = x.ID, Description = x.Description }).ToList();

            string html = "";
            foreach (GarmentDefectType garmentDefect in garmentDefectTypes)
            {
                html += "<tr><td idx='" + garmentDefect.ID + "' class='DefectCategoryTD' style='height: 8vh; text-align: center; vertical-align: middle;'><p>" + garmentDefect.Description + "</p></td></tr>";
            }

            return Content(html);
        }

        [HttpPost]
        public ActionResult GetDefectCode(string DefectTypeID)
        {
            List<GarmentDefectCode> garmentDefectCodes = _InspectionService.GetGarmentDefectCode(new GarmentDefectCode { GarmentDefectTypeID = DefectTypeID } ).ToList();

            string html = "";

            for (int i = 0; i <= garmentDefectCodes.Count() - 1; i += 2)
            {
                html += "<tr>";
                if (i <= garmentDefectCodes.Count() - 1) html += "<td idx='" + garmentDefectCodes[i].ID + "' class='DefectTypeID' style='height: 8vh; text-align: center; vertical-align: middle;'><p style='width: 9vw; word-wrap: break-word; padding: 0;'>" + garmentDefectCodes[i].Description + "</p></td>";
                if (i + 1 <= garmentDefectCodes.Count() - 1) html += "<td idx='" + garmentDefectCodes[i + 1].ID + "' class='DefectTypeTD' style='height: 8vh; text-align: center; vertical-align: middle;'><p style='width: 9vw; word-wrap: break-word; padding: 0;'>" + garmentDefectCodes[i + 1].Description + "</p></td>";
                html += "</tr>";

            }

            return Content(html);
        }

        [HttpPost]
        public JsonResult GetArea(string Type, string Location)
        {            
            List<Area> areas = _InspectionService.GetArea(new Area() 
                                { 
                                    Type = Type ,
                                    T = Location.Equals("T") ,
                                    B = Location.Equals("B") ,
                                    I = Location.Equals("I") ,
                                    O = Location.Equals("O") ,
                                })
                                .ToList();
            var result = areas.OrderBy(x => x.Seq).Select(x => new { x.Code }).ToList();
            return Json(result);
        }

        [HttpPost]
        public JsonResult GetDropDownList(string Type, string defectType)
        {
            List<DropDownList> downLists = _InspectionService.GetDropDownList(new DropDownList() { Type = Type, }).ToList();

            var result = downLists.OrderBy(x => x.Seq)
                    .Select(x => new 
                    { 
                        Description = defectType.Equals(InspectionService.DefectType.BAAuditCriteria) ? x.ID + " " + x.Description : x.Description 
                    })
                    .ToList();
            return Json(result);
        }
    }
}