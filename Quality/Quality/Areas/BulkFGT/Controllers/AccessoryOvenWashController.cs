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

        #region AccessoryOvenWash頁面
        public ActionResult Index()
        {
            Accessory_ViewModel accessory_ViewModel = new Accessory_ViewModel()
            {
                DataList = new List<Accessory_Result>(),
            };
            return View(accessory_ViewModel);
        }

        [HttpGet]
        public ActionResult IndexBack(string OrderID)
        {
            Accessory_ViewModel accessory_ViewModel = new Accessory_ViewModel()
            {
                OrderID = OrderID,
                StyleID = "",
                SeasonID = "",
                EarliestDate = System.DateTime.Now,
                EarliestSCIDel = System.DateTime.Now,
                TargetLeadTime = System.DateTime.Now,
                CompletionDate = System.DateTime.Now,
                ArticlePercent = 1,
                MtlCmplt = "",
                Remark = "",
                CreateBy = "",
                EditBy = "",

                DataList = new List<Accessory_Result>(),
            };

            Accessory_Result accessory_Result = new Accessory_Result()
            {
                AIR_LaboratoryID = "aaa",
                Seq1 = "01",
                Seq2 = "01",
                Seq = "",

                WKNo = "",
                WhseArrival = System.DateTime.Now,
                SCIRefno = "",
                Refno = "",
                SuppID = "",
                Supplier = "",
                Color = "",
                Size = "",
                ArriveQty = 123,
                InspDeadline = System.DateTime.Now,
                Result = 22,

                NonOven = false,
                OvenResult = "aa",
                OvenScale = "bb",
                OvenDate = System.DateTime.Now,
                OvenInspector = "cc",
                OvenRemark = "dd",

                NonWash = true,
                WashResult = "gg",
                WashScale = "bba",
                WashDate = System.DateTime.Now,
                WashInspector = "ttt",
                WashRemark = "a",
                Receiving = "b",
            };

            accessory_ViewModel.DataList.Add(accessory_Result);
            return View("Index", accessory_ViewModel);
        }


        [HttpPost]
        public ActionResult Index(string OrderID)
        {
            Accessory_ViewModel accessory_ViewModel = new Accessory_ViewModel()
            {
                OrderID = OrderID,
                StyleID = "",
                SeasonID = "",
                EarliestDate = System.DateTime.Now,
                EarliestSCIDel = System.DateTime.Now,
                TargetLeadTime = System.DateTime.Now,
                CompletionDate = System.DateTime.Now,
                ArticlePercent = 1,
                MtlCmplt = "",
                Remark = "",
                CreateBy = "",
                EditBy = "",

                DataList = new List<Accessory_Result>(),
            };

            Accessory_Result accessory_Result = new Accessory_Result()
            {
                AIR_LaboratoryID = "abc",
                Seq1 = "01",
                Seq2 = "02",
                Seq = "",

                WKNo = "",
                WhseArrival = System.DateTime.Now,
                SCIRefno = "",
                Refno = "",
                SuppID = "",
                Supplier = "",
                Color = "",
                Size = "",
                ArriveQty = 123,
                InspDeadline = System.DateTime.Now,
                Result = 22,

                NonOven = false,
                OvenResult = "aa",
                OvenScale = "bb",
                OvenDate = System.DateTime.Now,
                OvenInspector = "cc",
                OvenRemark = "dd",

                NonWash = true,
                WashResult = "gg",
                WashScale = "bba",
                WashDate = System.DateTime.Now,
                WashInspector = "ttt",
                WashRemark = "a",
                Receiving = "b",
            };

            accessory_ViewModel.DataList.Add(accessory_Result);
            return View(accessory_ViewModel);
        }

        [HttpPost]
        public ActionResult AccessorySave(Accessory_ViewModel Req)
        {
            return View();
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