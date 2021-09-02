using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using DatabaseObject.ResultModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class FabricOvenTestController : BaseController
    {
        private IFabricOvenTestService _FabricOvenTestService;
        private List<string> Results = new List<string>() { "Pass", "Fail" };
        private List<string> Temperatures = new List<string>() { "70", "90" };
        private List<string> Times = new List<string>() { "4", "24","48" };
        public FabricOvenTestController()
        {
            _FabricOvenTestService = new FabricOvenTestService();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.FabricOvenTest,,";
        }

        // GET: BulkFGT/FabricOvenTest
        public ActionResult Index()
        {

            FabricOvenTest_Result model = new FabricOvenTest_Result()
            {
                Main = new FabricOvenTest_Main(),
                Details = new List<FabricOvenTest_Detail>()
            };

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(string POID)
        {

            FabricOvenTest_Result model = _FabricOvenTestService.GetFabricOvenTest_Result("21051739BB");

            return View(model);
        }

        public ActionResult Detail()
        {

            FabricOvenTest_Detail_Result model = _FabricOvenTestService.GetFabricOvenTest_Detail_Result("21051739BB","");

            List<SelectListItem> ScaleIDList = new SetListItem().ItemListBinding(model.ScaleIDs);
            List<SelectListItem> ResultChangeList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> ResultStainList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> TimeList = new SetListItem().ItemListBinding(Times);

            ViewBag.ChangeScaleList = ScaleIDList;
            ViewBag.ResultChangeList = ResultChangeList;
            ViewBag.StainingScaleList = ScaleIDList;
            ViewBag.ResultStainList = ResultStainList;
            ViewBag.TemperatureList = TemperatureList;
            ViewBag.TimeList = TimeList;
            return View(model);
    }
}
}