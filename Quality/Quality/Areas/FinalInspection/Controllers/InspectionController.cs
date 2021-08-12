using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.FinalInspection.Controllers
{
    public class InspectionController : Controller
    {
        // GET: FinalInspection/Inspection
        public ActionResult Index()
        {
            return View();
        }
    }
}