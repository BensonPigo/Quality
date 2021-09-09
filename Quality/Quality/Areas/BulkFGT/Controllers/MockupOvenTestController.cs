using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
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
    }
}