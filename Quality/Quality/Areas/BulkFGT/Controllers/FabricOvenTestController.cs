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
        public FabricOvenTestController()
        {
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.FabricOvenTest,,";
        }

        // GET: BulkFGT/FabricOvenTest
        public ActionResult Index()
        {
            return View();
        }
    }
}