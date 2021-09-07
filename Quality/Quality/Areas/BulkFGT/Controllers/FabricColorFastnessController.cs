using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
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


        public ActionResult IndexBack(string POID)
        {
            FabricColorFastness_ViewModel model = _FabricColorFastness_Service.Get_Main("21051739BB");

            ViewBag.POID = POID;
            return View("Index", model);
        }


        public ActionResult Detail(string POID, string TestNo, string EditMode)
        {
            FabricColorFastness_ViewModel FabricColorFastnessModel = new FabricColorFastness_ViewModel();

            //ColorFastness_Result modelMaster = _FabricColorFastness_Service.GetDetailHeader("VM2CF21030110");
            //IList<Fabirc_ColorFastness_Detail_ViewModel> modelDetail = _FabricColorFastness_Service.GetDetailBody("VM2CF21030110");

            //Fabirc_ColorFastness_Detail_ViewModel model = new Fabirc_ColorFastness_Detail_ViewModel()
            //{
            //    Main = modelMaster,
            //};
            Fabric_ColorFastness_Detail_ViewModel model = new Fabric_ColorFastness_Detail_ViewModel();

            ColorFastness_Result modelMaster = new ColorFastness_Result();
            modelMaster.POID = "21051739BB";
            modelMaster.TestNo = 2;
            model.Main = modelMaster;
            ViewBag.Temperature_List = FabricColorFastnessModel.Temperature_List;
            ViewBag.Cycle_List = FabricColorFastnessModel.Cycle_List;
            ViewBag.Detergent_List = FabricColorFastnessModel.Detergent_List;
            ViewBag.Machine_List = FabricColorFastnessModel.Machine_List;
            ViewBag.Drying_List = FabricColorFastnessModel.Drying_List;
            ViewBag.EditMode = EditMode;
            ViewBag.FactoryID = this.FactoryID;
            return View(model);
        }


        //public ActionResult Detail()
        //{
        //    return View();
        //}
    }
}