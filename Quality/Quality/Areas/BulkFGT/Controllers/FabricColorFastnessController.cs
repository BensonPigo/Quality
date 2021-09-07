using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ViewModel.BulkFGT;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class FabricColorFastnessController : BaseController
    {
        private IFabricColorFastness_Service _FabricColorFastness_Service;
        // GET: BulkFGT/FabricColorFastness

        public FabricColorFastnessController()
        {
            _FabricColorFastness_Service = new FabricColorFastness_Service();

        }


        public ActionResult Index()
        {

            FabricColorFastness_ViewModel model = new FabricColorFastness_ViewModel()
            {
                ColorFastness_MainList = new List<ColorFastness_Result>()
            };

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(string POID)
        {

            FabricColorFastness_ViewModel model = _FabricColorFastness_Service.Get_Main("21051739BB");

            ViewBag.POID = POID;
            return View(model);
        }

        public ActionResult Detail()
        {
            return View();
        }
    }
}