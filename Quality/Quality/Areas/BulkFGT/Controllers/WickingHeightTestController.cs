using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Microsoft.Office.Interop.Excel;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class WickingHeightTestController : BaseController
    {
        private WickingHeightTestService _Service;
        public WickingHeightTestController()
        {
            _Service = new WickingHeightTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.WickingHeightTest,,";
        }

        // GET: BulkFGT/WickingHeightTest
        public ActionResult Index()
        {
            WickingHeightTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["WickingHeightTestModel"] != null)
            {
                model = (WickingHeightTest_ViewModel)TempData["WickingHeightTestModel"];
            }

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            WickingHeightTest_Request request = new WickingHeightTest_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            WickingHeightTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(WickingHeightTest_Request request)
        {
            WickingHeightTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult New()
        {
            WickingHeightTest_ViewModel model = _Service.GetDefaultModel(true);
                
            model.Main.EditType = "New";

            return View("Index", model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(WickingHeightTest_ViewModel requestModel)
        {
            WickingHeightTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["WickingHeightTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(WickingHeightTest_ViewModel requestModel)
        {
            WickingHeightTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["WickingHeightTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(WickingHeightTest_ViewModel requestModel)
        {
            WickingHeightTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }

            TempData["WickingHeightTestModel"] = model;

            return RedirectToAction("Index");
        }

        public ActionResult OrderIDCheck(string orderID)
        {
            WickingHeightTest_ViewModel model = _Service.GetOrderInfo(orderID);
            string ErrMsg = model.Result ? string.Empty : model.ErrorMessage;
            return Json(new { ErrMsg = ErrMsg, Result = model.Result, Main = model.Main });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode(string ReportNo, string Result)
        {
            CheckSession();
            WickingHeightTest_ViewModel result = _Service.EncodeAmend(new WickingHeightTest_ViewModel()
            {
                Main = new WickingHeightTest_Main()
                {
                    ReportNo = ReportNo,
                    Status = "Confirmed",
                    Result = Result
                },

            }, this.UserID);

            WickingHeightTest_ViewModel model = _Service.GetData(new WickingHeightTest_Request() { ReportNo = ReportNo });

            if (model.Main.Result == "Fail")
            {
                return Json(new { result.Result, ErrMsg = result.ErrorMessage, Action = "FailMail()" });
            }
            else
            {
                return Json(new { result.Result, ErrMsg = result.ErrorMessage, Action = "" });
            }
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Amend(string ReportNo, string Result)
        {
            CheckSession();
            WickingHeightTest_ViewModel result = _Service.EncodeAmend(new WickingHeightTest_ViewModel()
            {
                Main = new WickingHeightTest_Main()
                {
                    ReportNo = ReportNo,
                    Status = "New",
                    Result = Result
                },

            }, this.UserID);


            return Json(new { result.Result, ErrMsg = result.ErrorMessage });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ToExcel(string ReportNo)
        {
            WickingHeightTest_ViewModel result = _Service.GetReport(ReportNo, false);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + result.TempFileName;

            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ToPDF(string ReportNo)
        {
            WickingHeightTest_ViewModel result = _Service.GetReport(ReportNo, true);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + result.TempFileName;

            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }

        public JsonResult SendMailToMR(string ReportNo)
        {
            this.CheckSession();
            WickingHeightTest_ViewModel result = _Service.GetReport(ReportNo, false);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + result.TempFileName;

            return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, FileName = result.TempFileName });
        }
        public JsonResult SendMail(string ReportNo, string TO, string CC)
        {
            SendMail_Result result = _Service.SendMail(ReportNo, TO, CC);
            return Json(result);
        }
    }
}