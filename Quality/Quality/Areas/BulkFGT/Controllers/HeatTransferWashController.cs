using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class HeatTransferWashController : BaseController
    {
        private HeatTransferWashService _Service;
        public HeatTransferWashController()
        {
            _Service = new HeatTransferWashService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.HeatTransferWash,,";
        }

        // GET: BulkFGT/HeatTransferWash
        public ActionResult Index()
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Request = new HeatTransferWash_Request(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            return View(model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(HeatTransferWash_Request Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Request = Req,
                Main = new HeatTransferWash_Result(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            model = _Service.GetHeatTransferWash(Req);

            if (model.Result && !model.ReportNo_Source.Any())
            {
                model.ErrorMessage = "Data not found.";
            }
            else if (!model.Result)
            {
                string ErrorMessage = model.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
            }

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult NewSave(HeatTransferWash_ViewModel Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Request = new HeatTransferWash_Request(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };


            if (Req.Details == null)
            {
                Req.Details = new List<HeatTransferWash_Detail_Result>();
            }


            Req.Main.TestBeforePicture = Req.Main.TestBeforePicture == null ? null : ImageHelper.ImageCompress(Req.Main.TestBeforePicture);
            Req.Main.TestAfterPicture = Req.Main.TestAfterPicture == null ? null : ImageHelper.ImageCompress(Req.Main.TestAfterPicture);
            Req.Main.AddName = this.UserID;

            BaseResult result = _Service.Create(Req, this.MDivisionID, this.UserID, out string NewReportNo);

            if (!result.Result)
            {
                string ErrorMessage = model.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
            }
            else if (result.Result)
            {
                model = _Service.GetHeatTransferWash(new HeatTransferWash_Request()
                {
                    ReportNo = NewReportNo,
                });
            }

            return View("Index", model);
        }


        [SessionAuthorizeAttribute]
        [HttpPost]
        public JsonResult SPBlur(string POID)
        {
            string BrandID = string.Empty;
            string SeasonID = string.Empty;
            string StyleID = string.Empty;
            string Article = string.Empty;

            Orders order = new Orders();
            order.ID = POID;
            List<Orders> orderResult = _Service.GetOrders(order);

            if (orderResult.Count == 0)
            {
                return Json(new { ErrMsg = $"Cannot found SP# {POID}." });
            }

            if (orderResult.Count == 1)
            {
                BrandID = orderResult.FirstOrDefault().BrandID;
                SeasonID = orderResult.FirstOrDefault().SeasonID;
                StyleID = orderResult.FirstOrDefault().StyleID;

                Order_Qty order_qty = new Order_Qty();
                order_qty.ID = POID;
                List<Order_Qty> order_qtyResult = _Service.GetDistinctArticle(order_qty);

                if (order_qtyResult.Count == 1)
                {
                    Article = order_qtyResult.FirstOrDefault().Article;
                }

            }

            return Json(new { ErrMsg = "", BrandID = BrandID, SeasonID = SeasonID, StyleID = StyleID, Article = Article });
        }
    }
}