using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.FinalInspection.Controllers
{
    public class InspectionController : Controller
    {
        // Setting
        public ActionResult Index()
        {
            return View(new UserList());
        }

        public ActionResult Setting()
        {
            return View();
        }

        public ActionResult General()
        {
            return View();
        }

        public ActionResult CheckList()
        {
            return View();
        }

        public ActionResult AddDefect()
        {
            return View();
        }

        public ActionResult BeautifulProductAudit()
        {
            return View();
        }

        public ActionResult Moisture()
        {
            return View();
        }

        public ActionResult Measurement()
        {
            return View();
        }

        public ActionResult Others()
        {
            return View();
        }
        
    }
}