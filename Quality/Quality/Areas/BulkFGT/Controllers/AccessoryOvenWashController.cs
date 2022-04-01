using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject;
using DatabaseObject.ResultModel;
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
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.AccessoryOvenWash,,";
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

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string OrderID)
        {
            Accessory_ViewModel Req = new Accessory_ViewModel() { ReqOrderID = OrderID };

            Accessory_ViewModel model = new Accessory_ViewModel();

            model = _Service.GetMainData(Req);
            model.ReqOrderID = Req.ReqOrderID;

            return View("Index", model);
        }

        public ActionResult AccessoryOvenWash(string ReqOrderID)
        {
            Accessory_ViewModel model = new Accessory_ViewModel();
            model = _Service.GetMainData(new Accessory_ViewModel() { ReqOrderID = ReqOrderID });
            model.ReqOrderID = ReqOrderID;

            return View("Index", model);
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
        public ActionResult AccessorySave(Accessory_ViewModel Req)
        {
            Accessory_ViewModel model = new Accessory_ViewModel();
            Req.EditBy = this.UserID;
            _Service.Update(Req);
            _Service.UpdateInspPercent(Req.OrderID);
            Req.ReqOrderID = Req.OrderID;

            TempData["Req"] = Req;

            return Json(Req);
        }
        #endregion



        #region OvenTest頁面

        public ActionResult OvenTest(Accessory_Oven Req)
        {
            Accessory_Oven model = new Accessory_Oven();
            model = _Service.GetOvenTest(Req);
            ViewBag.FactoryID = this.FactoryID;

            return View(model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "OvenSave")]
        public ActionResult OvenSave(Accessory_Oven Req)
        {

            Accessory_Oven model = new Accessory_Oven();

            if (string.IsNullOrEmpty(Req.OvenInspector))
            {
                Req.OvenInspector = this.UserID;
            }
            Req.EditName = this.UserID;

            //修改
            model = _Service.UpdateOven(Req);

            if (!model.Result)
            {
                // 錯誤處理：新增一個Model承接ErrorMessage，並查詢原資料帶到畫面
                Accessory_Oven errorModel = new Accessory_Oven();
                errorModel = _Service.GetOvenTest(Req);
                errorModel.ErrorMessage = model.ErrorMessage;
                errorModel.ScaleData = model.ScaleData;

                return View("OvenTest", errorModel);

            }

            // 取得更新後MODEL
            model = _Service.GetOvenTest(Req);

            ViewBag.FactoryID = this.FactoryID;
            return View("OvenTest", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "OvenEncode")]
        public ActionResult OvenEncode(Accessory_Oven Req)
        {
            Accessory_Oven model = new Accessory_Oven();
            Req.EditName = this.UserID;
            Req.OvenEncode = true;

            //修改
            model = _Service.EncodeAmendOven(Req);

            _Service.UpdateInspPercent(Req.POID);
            if (!model.Result)
            {
                // 錯誤處理：新增一個Model承接ErrorMessage，並查詢原資料帶到畫面
                Accessory_Oven errorModel = new Accessory_Oven();
                errorModel = _Service.GetOvenTest(Req);
                errorModel.ErrorMessage = model.ErrorMessage;
                errorModel.ScaleData = model.ScaleData;

                return View("OvenTest", errorModel);

            }
            // 取得更新後MODEL
            model = _Service.GetOvenTest(Req);

            // Encode成功後，OvenResult是Fail則寄信
            // ISP20220193 規則修改
            if (model.OvenResult == "Fail")
            {
                model.ErrorMessage = "FailMail();";
            }

            ViewBag.FactoryID = this.FactoryID;
            return View("OvenTest", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "OvenAmend")]
        public ActionResult OvenAmend(Accessory_Oven Req)
        {
            Accessory_Oven model = new Accessory_Oven();
            Req.EditName = this.UserID;
            Req.OvenEncode = false;

            //修改
            model = _Service.EncodeAmendOven(Req);
            _Service.UpdateInspPercent(Req.POID);

            if (!model.Result)
            {
                // 錯誤處理：新增一個Model承接ErrorMessage，並查詢原資料帶到畫面
                Accessory_Oven errorModel = new Accessory_Oven();
                errorModel = _Service.GetOvenTest(Req);
                errorModel.ErrorMessage = model.ErrorMessage;
                errorModel.ScaleData = model.ScaleData;

                return View("OvenTest", errorModel);

            }

            // 取得更新後MODEL
            model = _Service.GetOvenTest(Req);

            ViewBag.FactoryID = this.FactoryID;
            return View("OvenTest", model);
        }

        [HttpPost]
        public JsonResult OvenFailMail(Accessory_Oven Req)
        {
            SendMail_Result result = _Service.SendOvenMail(Req);
            return Json(result);
        }

        [HttpPost]
        public JsonResult OvenReport(string AIR_LaboratoryID, string POID, string Seq1, string Seq2, bool IsToPDF)
        {
            BaseResult result = null;
            string FileName = string.Empty;
            
            result = _Service.OvenTestExcel(AIR_LaboratoryID, POID, Seq1, Seq2, IsToPDF, out FileName);

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        #endregion

        #region WashTest頁面

        public ActionResult WashTest(Accessory_Wash Req)
        {
            Accessory_Wash model = new Accessory_Wash();
            model = _Service.GetWashTest(Req);
            ViewBag.FactoryID = this.FactoryID;

            return View(model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashSave")]
        public ActionResult WashSave(Accessory_Wash Req)
        {

            Accessory_Wash model = new Accessory_Wash();

            if (string.IsNullOrEmpty(Req.WashInspector))
            {
                Req.WashInspector = this.UserID;
            }
            Req.EditName = this.UserID;

            //修改
            model = _Service.UpdateWash(Req);

            if (!model.Result)
            {
                // 錯誤處理：新增一個Model承接ErrorMessage，並查詢原資料帶到畫面
                Accessory_Wash errorModel = new Accessory_Wash();
                errorModel = _Service.GetWashTest(Req);
                errorModel.ErrorMessage = model.ErrorMessage;
                errorModel.ScaleData = model.ScaleData;

                return View("WashTest", errorModel);

            }

            // 取得更新後MODEL
            model = _Service.GetWashTest(Req);

            ViewBag.FactoryID = this.FactoryID;
            return View("WashTest", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashEncode")]
        public ActionResult WashEncode(Accessory_Wash Req)
        {
            Accessory_Wash model = new Accessory_Wash();
            Req.EditName = this.UserID;
            Req.WashEncode = true;


            //修改
            model = _Service.EncodeAmendWash(Req);
            _Service.UpdateInspPercent(Req.POID);

            if (!model.Result)
            {
                // 錯誤處理：新增一個Model承接ErrorMessage，並查詢原資料帶到畫面
                Accessory_Wash errorModel = new Accessory_Wash();
                errorModel = _Service.GetWashTest(Req);
                errorModel.ErrorMessage = model.ErrorMessage;
                errorModel.ScaleData = model.ScaleData;

                return View("WashTest", errorModel);

            }

            // 取得更新後MODEL
            model = _Service.GetWashTest(Req);

            // Encode成功後，WashResult是Fail則寄信
            // ISP20220193 規則修改
            if (model.WashResult == "Fail")
            {
                model.ErrorMessage = "FailMail();";
            }

            ViewBag.FactoryID = this.FactoryID;
            return View("WashTest", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashAmend")]
        public ActionResult WashAmend(Accessory_Wash Req)
        {
            Accessory_Wash model = new Accessory_Wash();
            Req.EditName = this.UserID;
            Req.WashEncode = false;

            //修改
            model = _Service.EncodeAmendWash(Req);
            _Service.UpdateInspPercent(Req.POID);

            if (!model.Result)
            {
                // 錯誤處理：新增一個Model承接ErrorMessage，並查詢原資料帶到畫面
                Accessory_Wash errorModel = new Accessory_Wash();
                errorModel = _Service.GetWashTest(Req);
                errorModel.ErrorMessage = model.ErrorMessage;
                errorModel.ScaleData = model.ScaleData;

                return View("WashTest", errorModel);

            }

            // 取得更新後MODEL
            model = _Service.GetWashTest(Req);

            ViewBag.FactoryID = this.FactoryID;
            return View("WashTest", model);
        }

        [HttpPost]
        public JsonResult WashFailMail(Accessory_Wash Req)
        {
            SendMail_Result result = _Service.SendWashMail(Req);
            return Json(result);
        }

        [HttpPost]
        public JsonResult WashReport(string AIR_LaboratoryID, string POID, string Seq1, string Seq2, bool IsToPDF)
        {
            BaseResult result = null;
            string FileName = string.Empty;

            result = _Service.WashTestExcel(AIR_LaboratoryID, POID, Seq1, Seq2, IsToPDF, out FileName);

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        #endregion

        #region WashingFastness (501)頁面

        public ActionResult WashingFastness(Accessory_WashingFastness Req)
        {
            Accessory_WashingFastness model = new Accessory_WashingFastness();
            model = _Service.GetWashingFastness(Req);
            ViewBag.FactoryID = this.FactoryID;

            return View(model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashingFastnessSave")]
        public ActionResult WashingFastnessSave(Accessory_WashingFastness Req)
        {

            Accessory_WashingFastness model = new Accessory_WashingFastness();

            if (string.IsNullOrEmpty(Req.WashingFastnessInspector))
            {
                Req.WashingFastnessInspector = this.UserID;
            }
            Req.EditName = this.UserID;
            Req.WashingFastnessEncode = false;

            //修改
            model = _Service.UpdateWashingFastness(Req);

            if (!model.Result)
            {
                // 錯誤處理：新增一個Model承接ErrorMessage，並查詢原資料帶到畫面
                Accessory_WashingFastness errorModel = new Accessory_WashingFastness();
                errorModel = _Service.GetWashingFastness(Req);
                errorModel.ErrorMessage = model.ErrorMessage;
                errorModel.ScaleData = model.ScaleData;

                return View("WashingFastness", errorModel);

            }

            // 取得更新後MODEL
            model = _Service.GetWashingFastness(Req);

            ViewBag.FactoryID = this.FactoryID;
            return View("WashingFastness", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashingFastnessEncode")]
        public ActionResult WashingFastnessEncode(Accessory_WashingFastness Req)
        {
            Accessory_WashingFastness model = new Accessory_WashingFastness();
            Req.EditName = this.UserID;
            Req.WashingFastnessEncode = true;

            //修改
            model = _Service.EncodeAmendWashingFastness(Req);
            _Service.UpdateInspPercent(Req.POID);

            if (!model.Result)
            {
                // 錯誤處理：新增一個Model承接ErrorMessage，並查詢原資料帶到畫面
                Accessory_WashingFastness errorModel = new Accessory_WashingFastness();
                errorModel = _Service.GetWashingFastness(Req);
                errorModel.ErrorMessage = model.ErrorMessage;
                errorModel.ScaleData = model.ScaleData;

                return View("WashingFastness", errorModel);

            }

            // 取得更新後MODEL
            model = _Service.GetWashingFastness(Req);

            // Encode成功後，WashResult是Fail則寄信
            // ISP20220193 規則修改
            if (model.WashingFastnessResult == "Fail")
            {
                model.ErrorMessage = "FailMail();";
            }

            ViewBag.FactoryID = this.FactoryID;
            return View("WashingFastness", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "WashingFastnessAmend")]
        public ActionResult WashWashingFastnessAmend(Accessory_WashingFastness Req)
        {
            Accessory_WashingFastness model = new Accessory_WashingFastness();
            Req.EditName = this.UserID;
            Req.WashingFastnessEncode = false;

            //修改
            model = _Service.EncodeAmendWashingFastness(Req);
            _Service.UpdateInspPercent(Req.POID);

            if (!model.Result)
            {
                // 錯誤處理：新增一個Model承接ErrorMessage，並查詢原資料帶到畫面
                Accessory_WashingFastness errorModel = new Accessory_WashingFastness();
                errorModel = _Service.GetWashingFastness(Req);
                errorModel.ErrorMessage = model.ErrorMessage;
                errorModel.ScaleData = model.ScaleData;

                return View("WashingFastness", errorModel);

            }

            // 取得更新後MODEL
            model = _Service.GetWashingFastness(Req);

            ViewBag.FactoryID = this.FactoryID;
            return View("WashingFastness", model);
        }

        [HttpPost]
        public JsonResult WashingFastnessFailMail(Accessory_WashingFastness Req)
        {
            SendMail_Result result = _Service.SendWashingFastnessMail(Req);
            return Json(result);
        }

        [HttpPost]
        public JsonResult WashingFastnessReport(string AIR_LaboratoryID, string POID, string Seq1, string Seq2, bool IsToPDF)
        {
            BaseResult result = null;
            string FileName = string.Empty;

            result = _Service.WashingFastnessExcel(AIR_LaboratoryID, POID, Seq1, Seq2, IsToPDF, out FileName);

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        #endregion

    }
}