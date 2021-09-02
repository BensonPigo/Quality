using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class AccessoryOvenWashController : Controller
    {

        #region AccessoryOvenWash頁面
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Index(string OrderID)
        {
            return View();
        }

        public ActionResult AccessoryEdit()
        {
            return View();
        }

        public ActionResult AccessoryUndo()
        {
            return View();
        }

        public ActionResult AccessorySave()
        {
            return View();
        }
        #endregion



        #region OvenTest頁面

        public ActionResult OvenTest()
        {
            return View();
        }

        public ActionResult OvenEdit()
        {
            return View();
        }

        public ActionResult OvenUndo()
        {
            return View();
        }

        public ActionResult OvenSave()
        {
            return View();
        }
        public ActionResult OvenEncode()
        {
            return View();
        }
        public ActionResult OvenAmend()
        {
            return View();
        }
        #endregion



        #region WashTest頁面
        public ActionResult WashTest()
        {
            return View();
        }

        public ActionResult WashEdit()
        {
            return View();
        }

        public ActionResult WashUndo()
        {
            return View();
        }

        public ActionResult WashSave()
        {
            return View();
        }
        public ActionResult WashEncode()
        {
            return View();
        }
        public ActionResult WashAmend()
        {
            return View();
        }
        #endregion

    }
}