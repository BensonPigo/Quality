using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
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
        private AccessoryOvenWashService _Service;

        public AccessoryOvenWashController()
        {
            _Service = new AccessoryOvenWashService();
        }

        #region AccessoryOvenWash頁面
        public ActionResult Index()
        {
            Accessory_ViewModel accessory_ViewModel = new Accessory_ViewModel()
            {
                DataList = new List<Accessory_Result>(),
            };
            return View(accessory_ViewModel);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(Accessory_ViewModel Req)
        {
            Accessory_ViewModel model = new Accessory_ViewModel();
            if (TempData["Req"] != null)
            {
                Req = (Accessory_ViewModel)TempData["Req"];
            }
            model = _Service.GetMainData(Req);
            model.ReqOrderID = Req.ReqOrderID;

            return View("Index", model);
        }



        [HttpPost]
        //[MultipleButton(Name = "action", Argument = "AccessorySave")]
        public ActionResult AccessorySave(Accessory_ViewModel Req)
        {
            Accessory_ViewModel model = new Accessory_ViewModel();
            _Service.Update(Req);
            Req.ReqOrderID = Req.OrderID;
            //model = _Service.GetMainData(Req);

            TempData["Req"] = Req;

            return Json(Req);
        }
        #endregion



        #region OvenTest頁面

        public ActionResult OvenTest(Accessory_Oven Req)
        {
            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<string> tempResult = new List<string>() {
                 "a","b", "c"
            };
            ViewBag.ScaleData = new SetListItem().ItemListBinding(tempResult);

            return View(Req);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "OvenSave")]
        public ActionResult OvenSave(Accessory_Oven Req)
        {
            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<string> tempResult = new List<string>() {
                 "a","b", "c"
            };
            ViewBag.ScaleData = new SetListItem().ItemListBinding(tempResult);

            return View("OvenTest", Req);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "OvenEncode")]
        public ActionResult OvenEncode(Accessory_Oven Req)
        {
            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<string> tempResult = new List<string>() {
                 "a","b", "c"
            };
            ViewBag.ScaleData = new SetListItem().ItemListBinding(tempResult);

            return View("OvenTest", Req);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "OvenAmend")]
        public ActionResult OvenAmend(Accessory_Oven Req)
        {
            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<string> tempResult = new List<string>() {
                 "a","b", "c"
            };
            ViewBag.ScaleData = new SetListItem().ItemListBinding(tempResult);

            return View("OvenTest", Req);
        }
        #endregion



        #region WashTest頁面
        public ActionResult WashTest(Accessory_Wash Req)
        {
            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<string> tempResult = new List<string>() {
                 "a","b", "c"
            };
            ViewBag.ScaleData = new SetListItem().ItemListBinding(tempResult);

            return View(Req);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashSave")]
        public ActionResult WashSave(Accessory_Wash Req)
        {
            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<string> tempResult = new List<string>() {
                 "a","b", "c"
            };
            ViewBag.ScaleData = new SetListItem().ItemListBinding(tempResult);

            return View("WashTest", Req);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashEncode")]
        public ActionResult WashEncode(Accessory_Wash Req)
        {
            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<string> tempResult = new List<string>() {
                 "a","b", "c"
            };
            ViewBag.ScaleData = new SetListItem().ItemListBinding(tempResult);

            return View("WashTest", Req);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashAmend")]
        public ActionResult WashAmend(Accessory_Wash Req)
        {
            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<string> tempResult = new List<string>() {
                 "a","b", "c"
            };
            ViewBag.ScaleData = new SetListItem().ItemListBinding(tempResult);
            return View("WashTest", Req);
        }
        #endregion

    }
}