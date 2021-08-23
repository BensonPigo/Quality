using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class GarmentTestController : BaseController
    {
        private IGarmentTest_Service _GarmentTest_Service;
        public GarmentTestController()
        {
            _GarmentTest_Service = new GarmentTest_Service();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.GarmentTest,,";
        }

        // GET: BulkFGT/GarmentTest
        public ActionResult Index()
        {
            GarmentTest_Request Req = new GarmentTest_Request();

            GarmentTest_Result Result = new GarmentTest_Result();
            ViewBag.GarmentTestRequest = Req;
            return View(Result);
        }

        [HttpPost]
        public ActionResult Index(GarmentTest_Request Req)
        {
            GarmentTest_Result Result = new GarmentTest_Result();
            ViewBag.GarmentTestRequest = Req;
            return View(Result);
        }
    }
}