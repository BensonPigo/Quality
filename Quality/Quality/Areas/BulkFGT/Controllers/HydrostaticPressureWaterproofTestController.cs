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
    public class HydrostaticPressureWaterproofTestController : BaseController
    {
        private HydrostaticPressureWaterproofTestService _Service;
        public HydrostaticPressureWaterproofTestController()
        {
            _Service = new HydrostaticPressureWaterproofTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.HydrostaticPressureWaterproofTest,,";
        }
        // GET: BulkFGT/HydrostaticPressureWaterproofTest
        public ActionResult Index()
        {
            HydrostaticPressureWaterproofTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["NewSaveHydrostaticPressureWaterproofTestModel"] != null)
            {
                model = (HydrostaticPressureWaterproofTest_ViewModel)TempData["NewSaveHydrostaticPressureWaterproofTestModel"];
            }
            else if (TempData["EditSaveHydrostaticPressureWaterproofTestModel"] != null)
            {
                model = (HydrostaticPressureWaterproofTest_ViewModel)TempData["EditSaveHydrostaticPressureWaterproofTestModel"];
            }
            else if (TempData["DeleteHydrostaticPressureWaterproofTestModel"] != null)
            {
                model = (HydrostaticPressureWaterproofTest_ViewModel)TempData["DeleteHydrostaticPressureWaterproofTestModel"];
            }

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            HydrostaticPressureWaterproofTest_Request request = new HydrostaticPressureWaterproofTest_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            HydrostaticPressureWaterproofTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(HydrostaticPressureWaterproofTest_Request request)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult New()
        {
            HydrostaticPressureWaterproofTest_ViewModel model = _Service.GetDefaultModel(true);

            model.Main.EditType = "New";

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(HydrostaticPressureWaterproofTest_ViewModel requestModel)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSaveHydrostaticPressureWaterproofTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(HydrostaticPressureWaterproofTest_ViewModel requestModel)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSaveHydrostaticPressureWaterproofTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(HydrostaticPressureWaterproofTest_ViewModel requestModel)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }
            TempData["DeleteHydrostaticPressureWaterproofTestModel"] = model;
            return RedirectToAction("Index");
        }


        public ActionResult OrderIDCheck(string orderID)
        {
            HydrostaticPressureWaterproofTest_ViewModel model = _Service.GetOrderInfo(orderID);

            if (!model.Result)
            {
                return Json(new { ErrMsg = model.ErrorMessage, Result = model.Result });
            }
            else
            {
                List<string> Article_Source = new List<string>();
                foreach (var item in model.Article_Source)
                {
                    string selected = string.Empty;
                    if (item.Value == model.Main.Article)
                    {
                        selected = "selected";
                    }
                    Article_Source.Add($"<option {selected} value='{item.Value}'>{item.Text}</option>");
                }

                return Json(new { ErrMsg = string.Empty, Result = model.Result, Main = model.Main, ArticleSource = Article_Source });
            }
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode(string ReportNo, string Result)
        {
            CheckSession();
            HydrostaticPressureWaterproofTest_ViewModel result = _Service.EncodeAmend(new HydrostaticPressureWaterproofTest_ViewModel()
            {
                Main = new HydrostaticPressureWaterproofTest_Main()
                {
                    ReportNo = ReportNo,
                    Status = "Confirmed",
                    Result = Result
                },

            }, this.UserID);

            HydrostaticPressureWaterproofTest_ViewModel model = _Service.GetData(new HydrostaticPressureWaterproofTest_Request() { ReportNo = ReportNo });

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
            HydrostaticPressureWaterproofTest_ViewModel result = _Service.EncodeAmend(new HydrostaticPressureWaterproofTest_ViewModel()
            {
                Main = new HydrostaticPressureWaterproofTest_Main()
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

            HydrostaticPressureWaterproofTest_ViewModel result = _Service.GetReport(ReportNo, false);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = "/TMP/" + Uri.EscapeDataString(result.TempFileName);

            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ToPDF(string ReportNo)
        {
            HydrostaticPressureWaterproofTest_ViewModel result = _Service.GetReport(ReportNo, true);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = "/TMP/" + Uri.EscapeDataString(result.TempFileName);

            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        public JsonResult SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = _Service.SendMail(ReportNo, TO, CC, Subject, Body, Files);
            return Json(result);
        }
    }
}