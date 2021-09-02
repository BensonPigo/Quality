using DatabaseObject.ViewModel.BulkFGT;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class AccessoryOvenWashController : BaseController
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
        public ActionResult AccessorySave(Accessory_ViewModel Req)
        {
            return View();
        }
        #endregion



        #region OvenTest頁面

        public ActionResult OvenTest(Accessory_Oven Req)
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "OvenSave")]
        public ActionResult OvenSave(Accessory_Oven Req)
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "OvenEncode")]
        public ActionResult OvenEncode(Accessory_Oven Req)
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "OvenAmend")]
        public ActionResult OvenAmend()
        {
            return View();
        }
        #endregion



        #region WashTest頁面
        public ActionResult WashTest(Accessory_Wash Req)
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashSave")]
        public ActionResult WashSave(Accessory_Wash Req)
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashEncode")]
        public ActionResult WashEncode(Accessory_Wash Req)
        {
            return View();
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashAmend")]
        public ActionResult WashAmend(Accessory_Wash Req)
        {
            return View();
        }
        #endregion

    }
}