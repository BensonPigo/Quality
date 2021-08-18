using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;


namespace Quality.Areas.SampleRFT.Controllers
{
    public class PicturesDummyController : BaseController
    {

        // GET: SampleRFT/PicturesDummy
        public ActionResult Index()
        {
            this.CheckSession();
            RFT_PicDuringDummyFitting_ViewModel model = new RFT_PicDuringDummyFitting_ViewModel();

            TempData["Model"] = null;
            return View(model);
        }
    }
}