
using BusinessLogicLayer.Service.FinalInspection;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.FinalInspection.Controllers
{
    public class QueryController : Controller
    {
        private QueryService Service = new QueryService();

        // GET: FinalInspection/Query
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Detail()
        {
            return View();
        }

        public ActionResult DownLoad()
        {

            return View();
        }
    }
}