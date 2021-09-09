using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class MockupOvenTestController : BaseController
    {
        private IMockupOvenService _MockupOvenService;

        public MockupOvenTestController()
        {
            _MockupOvenService = new MockupOvenService();
        }

        // GET: BulkFGT/MockupOvenTest
        public ActionResult Index()
        {

            MockupOven_ViewModel model = new MockupOven_ViewModel()
            {
                MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                ReportNo_Source = new List<string>(),
                Request = new MockupOven_Request(),

            };

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source;
            return View(model);
        }


        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(MockupOven_ViewModel Req)
        {
            MockupOven_Request MockupOven = new MockupOven_Request()
            { ReportNo = "PHOV180800003" };

            var model = _MockupOvenService.GetMockupOven(MockupOven);
            model.Request = new MockupOven_Request();
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            return View("Index", model);
        }

        [HttpPost]
        public JsonResult SPBlur(string POID)
        {
            string BrandID = string.Empty;
            string SeasonID = string.Empty;
            string StyleID = string.Empty;
            string Article =string.Empty;

            Orders order = new Orders();
            order.ID = POID;
            List<Orders> orderResult = _MockupOvenService.GetOrders(order);

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
                List<Order_Qty> order_qtyResult = _MockupOvenService.GetDistinctArticle(order_qty);

                if (order_qtyResult.Count == 1)
                {
                    Article = order_qtyResult.FirstOrDefault().Article;
                }

            }

            return Json(new { ErrMsg = "", BrandID= BrandID, SeasonID= SeasonID, StyleID= StyleID, Article= Article });
        }
    }
}