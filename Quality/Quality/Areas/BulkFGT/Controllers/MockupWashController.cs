using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class MockupWashController : Controller
    {
        public MockupWashController()
        {

        }

        // GET: BulkFGT/MockupWash
        public ActionResult Index()
        {
            return View();
        }
    }
}