using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

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

        [HttpPost]
        public ActionResult Query(Accessory_ViewModel Req)
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "AccessorySave")]
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