using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
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
        private List<string> MtlTypeIDs = new List<string>() { "KNIT", "WOVEN" };
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

            GarmentTest_Result Result = new GarmentTest_Result()
            {
                SizeCodes = new List<string>(),
                garmentTest = new GarmentTest_ViewModel()
                {
                    StyleID = "aaaa",
                    BrandID = "bbbb",
                },
                garmentTest_Details = new List<GarmentTest_Detail_ViewModel>() 
                {
                    new GarmentTest_Detail_ViewModel() { No = 1, ID = 1 }, 
                },
            };

            List<SelectListItem> SizeCodeList = new SetListItem().ItemListBinding(Result.SizeCodes);
            List<SelectListItem> MtlTypeIDList = new SetListItem().ItemListBinding(this.MtlTypeIDs);

            ViewBag.SizeCodeList = SizeCodeList;
            ViewBag.MtlTypeIDList = MtlTypeIDList;
            
            ViewBag.GarmentTestRequest = Req;
            return View(Result);
        }

        [HttpPost]
        public ActionResult Index(GarmentTest_Request Req)
        {
            GarmentTest_Result Result = new GarmentTest_Result()
            {
                SizeCodes = new List<string>(),
                garmentTest = new GarmentTest_ViewModel(),
                garmentTest_Details = new List<GarmentTest_Detail_ViewModel>()
                {
                    new GarmentTest_Detail_ViewModel() { No = 1, ID = 1 },
                },
            };

            List<SelectListItem> SizeCodeList = new SetListItem().ItemListBinding(Result.SizeCodes);
            List<SelectListItem> MtlTypeIDList = new SetListItem().ItemListBinding(this.MtlTypeIDs);
            ViewBag.SizeCodeList = SizeCodeList;
            ViewBag.MtlTypeIDList = MtlTypeIDList;
            ViewBag.GarmentTestRequest = Req;
            return View(Result);
        }

        [HttpPost]
        public JsonResult SaveDetail(List<GarmentTest_Detail> details)
        {
            var result = new { Result = false, ErrMsg = "Err" };
            return Json(result);
        }


        public ActionResult Detail(string ID, string No)
        {
            return View();
        }
    }
}