using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ToolKit;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class InspectionController : BaseController
    {
        private IInspectionService _InspectionService;
        public InspectionController()
        {
            _InspectionService = new InspectionService();
            if (this.SelectItemData == null || this.SelectItemData.Count() == 0)
                this.SelectItemData = _InspectionService.GetSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, Brand = this.Brand }).ToList();

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
                Brand = this.Brand,
            };

            List<SelectListItem> FactoryitemList = new SetListItem().ItemListBinding(this.Factorys);
            List<SelectListItem> LineList = new SetListItem().ItemListBinding(this.Lines);

            ViewBag.FactoryList = FactoryitemList;
            ViewBag.LineList = LineList;
            ViewBag.BrandList = this.Brands;
            ViewBag.UserID = this.UserID;
            ViewBag.ShowSwitch = 1;

            return View(setting);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult Index(Inspection_ViewModel request)
        {
            this.CheckSession();
            this.FactoryID = string.IsNullOrEmpty(request.FactoryID) ? this.FactoryID : request.FactoryID;
            this.Line = string.IsNullOrEmpty(request.Line) ? this.Line : request.Line;
            this.WorkDate = request.InspectionDate = CheckWorkDate.Check(request.InspectionDate);
            this.Brand = request.Brand;
            this.SelectItemData = _InspectionService.GetSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, Brand = this.Brand }).ToList();
            List<SelectListItem> FactoryitemList = new SetListItem().ItemListBinding(this.Factorys);
            List<SelectListItem> LineList = new SetListItem().ItemListBinding(this.Lines);

            ViewBag.FactoryList = FactoryitemList;
            ViewBag.LineList = LineList;
            ViewBag.BrandList = this.Brands;
            ViewBag.UserID = this.UserID;
            ViewBag.ShowSwitch = 0;

            return View(request);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult CheckBrand(string Brand)
        {
            this.CheckSession();
            Inspection_ViewModel viewModel = new Inspection_ViewModel();

            if (string.IsNullOrEmpty(Brand))
            {
                viewModel.Result = false;
                viewModel.Brand = string.Empty;
                return Json(viewModel);
            }

            var query = this.Brands.Where(x => x.Equals(HttpUtility.HtmlDecode(Brand))).ToList();

            viewModel.Result = query.Any();
            viewModel.Brand = query.Any() ? query.FirstOrDefault() : this.Brand;
            return Json(viewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetStyle(string OrderID)
        {
            this.CheckSession();

            List<string> styles = this.SelectItemData.Where(x => string.IsNullOrEmpty(OrderID) || x.OrderID.Equals(OrderID))
                                    .GroupBy(x => x.StyleID)
                                    .Select(x => "**" + x.Key + "@@").ToList();

            string html = string.Join(",", styles);

            html = html.Replace(",", string.Empty).Replace("**", "<li><a href='#'>").Replace("@@", "</a></li>");
            //string html = "";
            //foreach (string style in styles)
            //{
            //    html += "<li><a href='#'>" + style + "</a></li>";
            //}

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult CheckStyle(string StyleID)
        {
            this.CheckSession();
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(StyleID))
            {
                return Json(viewModel);
            }

            var query = _InspectionService
                            .CheckSelectItemData(new Inspection_ViewModel() 
                                                { 
                                                    FactoryID = this.FactoryID, 
                                                    StyleID = StyleID ,
                                                    Brand = this.Brand,
                                                }, InspectionService.SelectType.StyleID)
                            .ToList();

            viewModel.Result = query.Any();
            viewModel.StyleID = query.Any() ? query.FirstOrDefault().StyleID : string.Empty;
            return Json(viewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetSP()
        {
            this.CheckSession();
            List<string> orderIDs = SelectItemData
                                    .GroupBy(x => x.OrderID)
                                    .Select(x => "**" + x.Key + "@@").ToList();

            string html = string.Join(",", orderIDs);

            html = html.Replace(",", string.Empty).Replace("**", "<li><a href='#'>").Replace("@@", "</a></li>");
            //string html = "";
            //foreach (string orderID in orderIDs)
            //{
            //    html += "<li><a href='#'>" + orderID + "</a></li>";
            //}

            return Content(html);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult FilterSP(string SPno)
        {
            this.CheckSession();
            if (string.IsNullOrEmpty(SPno))
            {
                return Content(string.Empty);
            }

            List<string> orderIDs = SelectItemData
                                    .Where(o=>o.OrderID.StartsWith(SPno))
                                    .GroupBy(x => x.OrderID)
                                    .Select(x => "**" + x.Key + "@@").ToList();

            string html = string.Join(",", orderIDs);

            html = html.Replace(",", string.Empty).Replace("**", "<li><a href='#'>").Replace("@@", "</a></li>");
            //string html = "";
            //foreach (string orderID in orderIDs)
            //{
            //    html += "<li><a href='#'>" + orderID + "</a></li>";
            //}

            return Content(html);
        }


        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult CheckSP(string OrderID)
        {
            this.CheckSession();
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(OrderID))
            {
                return Json(viewModel);
            }

            List<Inspection_ViewModel> result = _InspectionService.CheckSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, OrderID = OrderID, Brand = this.Brand, }
                                                                                    , InspectionService.SelectType.OrderID)
                                                                   .ToList();
            viewModel = result.Any() ? result.FirstOrDefault() : viewModel;
            viewModel.Result = result.Any();

            if (!viewModel.Result)
            {
                List<Inspection_ChkOrderID_ViewModel> result2 = _InspectionService.CheckSelectItemData_SP(new Inspection_ViewModel() { FactoryID = this.FactoryID, OrderID = OrderID, Brand = this.Brand, }).ToList();
                string ErrMsg = "SP# not found";
                if (result2.Count > 0)
                {
                    if (result2.FirstOrDefault().Inpsected) ErrMsg = "Already inpsected!";
                    //if (result2.FirstOrDefault().PulloutComplete) ErrMsg = "Already pulled out!";
                }
                viewModel.ErrMsg = ErrMsg;
            }

            return Json(viewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetArticle(string StyleID, string OrderID)
        {
            this.CheckSession();
            if (string.IsNullOrEmpty(StyleID) && string.IsNullOrEmpty(OrderID))
            {
                return Content(string.Empty);
            }
            List<string> articles = _InspectionService
                                    .GetSelectItemData(new Inspection_ViewModel() { StyleID = StyleID, OrderID = OrderID })
                                    .GroupBy(x => x.Article)
                                    .Select(x => "**" + x.Key + "@@").ToList();

            string html = string.Join(",", articles);

            html = html.Replace(",", string.Empty).Replace("**", "<li><a href='#'>").Replace("@@", "</a></li>");
            //string html = "";
            //foreach (string article in articles)
            //{
            //    html += "<li><a href='#'>" + article + "</a></li>";
            //}

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult CheckArticle(string StyleID, string OrderID, string Article)
        {
            this.CheckSession();
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(Article))
            {
                return Json(viewModel);
            }

            List<Inspection_ViewModel> result = _InspectionService.CheckSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, StyleID = StyleID, OrderID = OrderID, Article = Article, Brand = this.Brand, }
                                                                                    , InspectionService.SelectType.Article)
                                                                   .ToList();
            viewModel = result.Any() ? result.FirstOrDefault() : viewModel;
            viewModel.Result = result.Any();
            return Json(viewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetSize(string StyleID, string OrderID, string Article)
        {
            this.CheckSession();
            if (string.IsNullOrEmpty(StyleID) && string.IsNullOrEmpty(OrderID) && string.IsNullOrEmpty(Article))
            {
                return Content(string.Empty);
            }
            List<string> sizes = _InspectionService
                                    .GetSelectItemData(new Inspection_ViewModel() { StyleID = StyleID, OrderID = OrderID, Article = Article })
                                    .GroupBy(x => x.Size)
                                    .Select(x => "**" + x.Key + "@@").ToList();

            string html = string.Join(",", sizes);

            html = html.Replace(",", string.Empty).Replace("**", "<li><a href='#'>").Replace("@@", "</a></li>");

            //string html = "";
            //foreach (string size in sizes)
            //{
            //    html += "<li><a href='#'>" + size + "</a></li>";
            //}

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult CheckSize(string StyleID, string OrderID, string Article, string Size)
        {
            this.CheckSession();
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(Size))
            {
                return Json(viewModel);
            }

            List<Inspection_ViewModel> result = _InspectionService.CheckSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, StyleID = StyleID, OrderID = OrderID, Article = Article, Size = Size, Brand = this.Brand, }
                                                                                    , InspectionService.SelectType.Size)
                                                                   .ToList();
            viewModel = result.Any() ? result.FirstOrDefault() : viewModel;
            viewModel.Result = result.Any();
            return Json(viewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetProductType(string StyleID, string OrderID, string Article, string Size)
        {
            this.CheckSession();
            if (string.IsNullOrEmpty(StyleID) && string.IsNullOrEmpty(OrderID) && string.IsNullOrEmpty(Article) && string.IsNullOrEmpty(Size))
            {
                return Content(string.Empty);
            }
            List<string> productTypes = _InspectionService
                                    .GetSelectItemData(new Inspection_ViewModel() { StyleID = StyleID, OrderID = OrderID, Article = Article, Size = Size })
                                    .GroupBy(x => x.ProductType)
                                    .Select(x => "**" + x.Key + "@@").ToList();

            string html = string.Join(",", productTypes);

            html = html.Replace(",", string.Empty).Replace("**", "<li><a href='#'>").Replace("@@", "</a></li>");
            //string html = "";
            //foreach (string productType in productTypes)
            //{
            //    html += "<li><a href='#'>" + productType + "</a></li>";
            //}

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult CheckProductType(string StyleID, string OrderID, string Article, string Size, string ProductType)
        {
            this.CheckSession();
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(ProductType))
            {
                return Json(viewModel);
            }

            List<Inspection_ViewModel> result = _InspectionService.CheckSelectItemData(new Inspection_ViewModel() { FactoryID = this.FactoryID, StyleID = StyleID, OrderID = OrderID, Article = Article, Size = Size, ProductType = ProductType, Brand = this.Brand, }
                                                                                    , InspectionService.SelectType.ProductType)
                                                                   .ToList();
            viewModel = result.Any() ? result.FirstOrDefault() : viewModel;
            viewModel.Result = result.Any();
            return Json(viewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult RefreshQty(string StyleID, string OrderID, string Article, string Size, string ProductType)
        {
            this.CheckSession();
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
                if (result.Result)
                {
                    Inspection_ViewModel inspection = inspection_Views.FirstOrDefault();
                    result.SizeQty = inspection.SizeQty;
                    result.SizeBalanceQty = inspection.SizeBalanceQty;
                    result.OrderQty = inspection.OrderQty;
                    result.OrderBalanceQty = inspection.OrderBalanceQty;
                }
            }

            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Pass(string StyleID, string OrderID, string Article, string Size, string ProductType)
        {
            this.CheckSession();
            Inspection_ViewModel viewModel = new Inspection_ViewModel();

            List<Inspection_ChkOrderID_ViewModel> result2 = _InspectionService.CheckSelectItemData_SP(new Inspection_ViewModel() { FactoryID = this.FactoryID, OrderID = OrderID, Brand = this.Brand, }).ToList();
            if (result2.Any() && result2.FirstOrDefault().Inpsected)
            {
                viewModel.Result = false;
                viewModel.ErrMsg = "Already inpsected!";
            }
            else
            {
                InspectionSave_ViewModel inspectionSave_View = new InspectionSave_ViewModel()
                {
                    rft_Inspection = new RFT_Inspection(),
                    fT_Inspection_Details = new List<RFT_Inspection_Detail>(),
                };

                RFT_Inspection rFT_Inspection = new RFT_Inspection()
                {
                    OrderID = OrderID,
                    Article = Article,
                    Location = ProductType,
                    Size = Size,
                    Line = this.Line,
                    FactoryID = this.FactoryID,
                    Status = "Pass",
                    AddDate = DateTime.Now,
                    AddName = this.UserID,
                    InspectionDate = this.WorkDate,
                };

                inspectionSave_View.StyleID = StyleID;
                inspectionSave_View.rft_Inspection = rFT_Inspection;
                InspectionSave_ViewModel save_ViewModel = _InspectionService.SaveRFTInspection(inspectionSave_View);

                viewModel.Result = save_ViewModel.Result;
                viewModel.ErrMsg = save_ViewModel.ErrMsg;
            }

            return Json(viewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult ChkInspQty(string OrderID, string Size, string Article)
        {
            this.CheckSession();
            InspectionSave_ViewModel inspectionSave_View = new InspectionSave_ViewModel();
            inspectionSave_View.rft_Inspection = new RFT_Inspection();
            if (string.IsNullOrEmpty(OrderID) || string.IsNullOrEmpty(Size))
            {
                return Json(new { Result = false, ErrMsg = "OrderID, Size is Empty" }) ;
            }
            InspectionSave_ViewModel result = new InspectionSave_ViewModel();
            List <Inspection_ChkOrderID_ViewModel> result2 = _InspectionService.CheckSelectItemData_SP(new Inspection_ViewModel() { FactoryID = this.FactoryID, OrderID = OrderID, Brand = this.Brand, }).ToList();
            if (result2.Any() && result2.FirstOrDefault().Inpsected)
            {
                result.Result = false;
                result.ErrMsg = "Already inpsected!";
            }
            else
            {
                inspectionSave_View.rft_Inspection.OrderID = OrderID;
                inspectionSave_View.rft_Inspection.Size = Size;
                inspectionSave_View.rft_Inspection.Article = Article;

                result = _InspectionService.ChkInspQty(inspectionSave_View);
            }

            return Json(new { result.Result, result .ErrMsg} );
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetDefectType()
        {
            this.CheckSession();
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
        [SessionAuthorizeAttribute]
        public ActionResult GetDefectCode(string DefectTypeID)
        {
            this.CheckSession();
            List<GarmentDefectCode> garmentDefectCodes = _InspectionService.GetGarmentDefectCode(new GarmentDefectCode { GarmentDefectTypeID = DefectTypeID } ).ToList();

            string html = "";

            for (int i = 0; i <= garmentDefectCodes.Count() - 1; i += 2)
            {
                html += "<tr>";
                if (i <= garmentDefectCodes.Count() - 1) html += "<td idx='" + garmentDefectCodes[i].ID + "' class='DefectTypeID' style='height: 8vh; text-align: center; vertical-align: middle;'><p style='width: 9vw; word-wrap: break-word; padding: 0;'>" + garmentDefectCodes[i].Description + "</p></td>";
                if (i + 1 <= garmentDefectCodes.Count() - 1) html += "<td idx='" + garmentDefectCodes[i + 1].ID + "' class='DefectTypeID' style='height: 8vh; text-align: center; vertical-align: middle;'><p style='width: 9vw; word-wrap: break-word; padding: 0;'>" + garmentDefectCodes[i + 1].Description + "</p></td>";
                html += "</tr>";

            }

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetArea(string Type, string Location)
        {
            this.CheckSession();
            List<Area> areas = _InspectionService.GetArea(new Area() 
                                { 
                                    Type = Type ,
                                    T = Location.Equals("T") ,
                                    B = Location.Equals("B") ,
                                    I = Location.Equals("I") ,
                                    O = Location.Equals("O") ,
                                })
                                .ToList();

            string html = "";

            for (int i = 0; i <= areas.Count() - 1; i += 2)
            {
                html += "<tr>";
                if (i <= areas.Count() - 1) html += "<td idx='" + areas[i].Code + "' class='Area2ID' style='height: 8vh; text-align: center; vertical-align: middle;'><p style='width: 9vw; word-wrap: break-word; padding: 0;'>" + areas[i].Code + "</p></td>";
                if (i + 1 <= areas.Count() - 1) html += "<td idx='" + areas[i + 1].Code + "' class='Area2ID' style='height: 8vh; text-align: center; vertical-align: middle;'><p style='width: 9vw; word-wrap: break-word; padding: 0;'>" + areas[i + 1].Code + "</p></td>";
                html += "</tr>";

            }

            return Content(html); 
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetDropDownList(string Type, string defectType)
        {
            this.CheckSession();
            List<DropDownList> downLists = _InspectionService.GetDropDownList(new DropDownList() { Type = Type, }).ToList();


            string html = "";

            for (int i = 0; i <= downLists.Count() - 1; i++)
            {
                string text = defectType.Equals(InspectionService.DefectType.BAAuditCriteria) ? downLists[i].ID + " " + downLists[i].Description : downLists[i].Description ;
                html += "<tr>";
                html += "<td idx='" + downLists[i].ID + "' class='" + defectType + "ID' style='height: 8vh; text-align: center; vertical-align: middle;'><p style='width: 9vw; word-wrap: break-word; padding: 0;'>" + text + "</p></td>";
                html += "</tr>";
            }

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult GetReworkCard(string fixType)
        {
            this.CheckSession();
            if (string.IsNullOrEmpty(this.FactoryID) || string.IsNullOrEmpty(this.Line))
            {
                return Json(new { No = "", Status = "" });
            }

            ReworkCard rework = new ReworkCard()
            {
                FactoryID = this.FactoryID,
                Line = this.Line,
                Type = fixType,
            };

            List<ReworkCard> reworks = _InspectionService.GetReworkCards(rework).ToList();

            var result = reworks.Select(x => new { No = x.No, Status = x.Status }).ToList();
            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult ReworkCardSave(InspectionSave_ViewModel saveView)
        {
            this.CheckSession();
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            InspectionSave_ViewModel inspectionSave_View = new InspectionSave_ViewModel()
            {
                rft_Inspection = new RFT_Inspection(),
                fT_Inspection_Details = new List<RFT_Inspection_Detail>(),
            };

            inspectionSave_View.rft_Inspection = saveView.rft_Inspection;
            inspectionSave_View.rft_Inspection.ReworkCardType = saveView.rft_Inspection.FixType;
            inspectionSave_View.rft_Inspection.Line = this.Line;
            inspectionSave_View.rft_Inspection.FactoryID = this.FactoryID;
            inspectionSave_View.rft_Inspection.Status = "Reject";
            inspectionSave_View.rft_Inspection.AddName = this.UserID;
            inspectionSave_View.rft_Inspection.InspectionDate = this.WorkDate;
            inspectionSave_View.StyleID = saveView.StyleID;
            inspectionSave_View.fT_Inspection_Details = saveView.fT_Inspection_Details;

            InspectionSave_ViewModel save_ViewModel = _InspectionService.SaveRFTInspection(inspectionSave_View);
            viewModel.Result = save_ViewModel.Result;
            viewModel.ErrMsg = save_ViewModel.ErrMsg;
            return Json(viewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult GetReworkList()
        {
            this.CheckSession();
            List<ReworkList_ViewModel> reworkList_View = _InspectionService.GetReworkList(new ReworkList_ViewModel() 
            { 
                FactoryID = this.FactoryID, 
                Line = this.Line, 
            });
             
            return Json(reworkList_View);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ReworkListSaveAction(List<RFT_Inspection> reworkLists, InspectionService.ReworkListType type)
        {
            this.CheckSession();
            foreach (RFT_Inspection item in reworkLists)
            {                               
                item.EditName = this.UserID;
                item.InspectionDate = type.Equals(InspectionService.ReworkListType.Pass) ? this.WorkDate : (DateTime?)null;
            }

            InspectionSave_ViewModel result = _InspectionService.SaveReworkListAction(reworkLists, type);
            return Json(new { Result = result.Result, ErrMsg = result.ErrMsg });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ReworkListAddReject(RFT_Inspection_Detail detail)
        {
            this.CheckSession();
            InspectionSave_ViewModel result = _InspectionService.SaveReworkListAddReject(detail);
            return Json(new { Result = result.Result, ErrMsg = result.ErrMsg });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult SaveReworkListDelete(LogIn_Request logIn, List<RFT_Inspection> datas)
        {
            this.CheckSession();
            InspectionSave_ViewModel result = _InspectionService.SaveReworkListDelete(logIn, datas);
            return Json(new { Result = result.Result, ErrMsg = result.ErrMsg });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult DQSReasonGet()
        {
            this.CheckSession();
            List<DQSReason> dQSReasons = _InspectionService.GetDQSReason(new DQSReason() { Type = "DP", Junk = false }) ;
            string html = "";
            foreach (DQSReason reason in dQSReasons)
            {
                html += "<input idx='" + reason.ID + "' type='button' value='" + reason.Description + "'>";
            }

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult MeasurementGet(string OrderID, string SizeCode)
        {
            this.CheckSession();
            List<RFT_Inspection_Measurement_ViewModel> rFT_Inspection_Measurements = _InspectionService.GetMeasurement(OrderID, SizeCode, this.UserID);
            return Json(rFT_Inspection_Measurements);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult MeasurementSave(List<RFT_Inspection_Measurement> measurement)
        {
            this.CheckSession();
            if (string.IsNullOrEmpty(this.FactoryID) || string.IsNullOrEmpty(this.Line))
            {
                Session.Abandon();
                this.CheckSession();
            }

            RFT_Inspection_Measurement_ViewModel rtn = new RFT_Inspection_Measurement_ViewModel();
            if (measurement == null)
            {
                rtn.Result = false;
                rtn.ErrMsg = "Please input data.";

            }
            else
            {
                foreach (RFT_Inspection_Measurement item in measurement)
                {
                    item.FactoryID = this.FactoryID;
                    item.Line = this.Line;
                }
                RFT_Inspection_Measurement_ViewModel viewModel = _InspectionService.SaveMeasurement(measurement);
                rtn = viewModel;
            }


            return Json(new { rtn.Result, rtn.ErrMsg });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult CFTCommentsGet(string OrderID, string StyleID, string Season, string SampleStage)
        {
            this.CheckSession();
            List<RFT_OrderComments_ViewModel> viewModel = new List<RFT_OrderComments_ViewModel>();
            if (string.IsNullOrEmpty(OrderID))
            {
                return Json(viewModel);
            }

            viewModel = _InspectionService.GetRFT_OrderComments(new RFT_OrderComments()
                                        {
                                            OrderID = OrderID,
                                        });

            string html = "";
            foreach (RFT_OrderComments_ViewModel item in viewModel)
            {
                html += "<tr>";
                html += "<td><p>" + item.PMS_RFTCommentsDescription  + "</p></td>";
                html += "<td><textarea id='" + item.PMS_RFTCommentsDescription + item.PMS_RFTCommentsID + "' idx='" + item.PMS_RFTCommentsID  + "' cols='85'> " + item.Comnments + " </textarea></td>";
                html += "</tr>";
            }

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult CFTCommentsSave(List<RFT_OrderComments> duringDummyFitting)
        {
            this.CheckSession();
            RFT_OrderComments_ViewModel rFT_PicDuringDummyFitting_ViewModel = _InspectionService.SaveRFT_OrderComments(duringDummyFitting);

            return Json(rFT_PicDuringDummyFitting_ViewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult CFTCommentsSend(string OrderID)
        {
            this.CheckSession();
            RFT_OrderComments_ViewModel rFT_PicDuringDummyFitting_ViewModel = _InspectionService.SendMailRFT_OrderComments(new RFT_OrderComments { OrderID = OrderID }, this.UserID);
            return Json(rFT_PicDuringDummyFitting_ViewModel);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult PicturesGet(string OrderID, string StyleID, string Article, string Size)
        {
            this.CheckSession();
            Inspection_ViewModel viewModel = new Inspection_ViewModel();
            if (string.IsNullOrEmpty(OrderID))
            {
                return Json(viewModel);
            }

            RFT_PicDuringDummyFitting result = _InspectionService.GetRFT_PicDuringDummyFitting(new RFT_PicDuringDummyFitting() 
                                                { 
                                                    OrderID = OrderID, 
                                                    Article = Article, 
                                                    Size = Size 
                                                });
            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult PicturesSave(RFT_PicDuringDummyFitting duringDummyFitting)
        {
            this.CheckSession();
            RFT_PicDuringDummyFitting_ViewModel rFT_PicDuringDummyFitting_ViewModel = _InspectionService.SaveRFT_PicDuringDummyFitting(duringDummyFitting);

            return Json(rFT_PicDuringDummyFitting_ViewModel);
        }
    }
}