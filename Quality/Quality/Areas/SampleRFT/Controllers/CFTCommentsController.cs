
using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.SampleRFT;
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
    public class CFTCommentsController : BaseController
    {
        // GET: SampleRFT/CFTComments
        public ActionResult Index()
        {
            return View();
        }
    }
}