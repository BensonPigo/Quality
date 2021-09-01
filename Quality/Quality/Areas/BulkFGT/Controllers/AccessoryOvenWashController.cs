using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class AccessoryOvenWashController : Controller
    {
        // GET: BulkFGT/AccessoryOvenWash
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult OvenTest()
        {
            return View();
        }

        public ActionResult WashTest()
        {
            return View();
        }
    }
}