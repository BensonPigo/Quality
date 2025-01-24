using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.RequestModel;
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
using System.Web.Services.Description;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class RandomTumblePillingTestController : BaseController
    {
        private RandomTumblePillingTestService _Service;
        public RandomTumblePillingTestController()
        {
            _Service = new RandomTumblePillingTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.RandomTumblePillingTest,,";
        }
        // GET: BulkFGT/RandomTumblePillingTest
        public ActionResult Index()
        {
            RandomTumblePillingTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["NewSaveRandomTumblePillingTestModel"] != null)
            {
                model = (RandomTumblePillingTest_ViewModel)TempData["NewSaveRandomTumblePillingTestModel"];
            }
            else if (TempData["EditSaveRandomTumblePillingTestModel"] != null)
            {
                model = (RandomTumblePillingTest_ViewModel)TempData["EditSaveRandomTumblePillingTestModel"];
            }
            else if (TempData["DeleteRandomTumblePillingTestModel"] != null)
            {
                model = (RandomTumblePillingTest_ViewModel)TempData["DeleteRandomTumblePillingTestModel"];
            }

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            RandomTumblePillingTest_Request request = new RandomTumblePillingTest_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            RandomTumblePillingTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(RandomTumblePillingTest_Request request)
        {
            RandomTumblePillingTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult New()
        {
            RandomTumblePillingTest_ViewModel model = _Service.GetDefaultModel(true);

            model.Main.EditType = "New";

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(RandomTumblePillingTest_ViewModel requestModel)
        {
            RandomTumblePillingTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSaveRandomTumblePillingTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(RandomTumblePillingTest_ViewModel requestModel)
        {
            RandomTumblePillingTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSaveRandomTumblePillingTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(RandomTumblePillingTest_ViewModel requestModel)
        {
            RandomTumblePillingTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }
            TempData["DeleteRandomTumblePillingTestModel"] = model;
            return RedirectToAction("Index");
        }


        public ActionResult OrderIDCheck(string orderID)
        {
            RandomTumblePillingTest_ViewModel model = _Service.GetOrderInfo(orderID);

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


                return Json(new { ErrMsg = string.Empty, Result = model.Result, Main = model.Main, ArticleSource = Article_Source  });
            }
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode(string ReportNo, string Result)
        {
            CheckSession();
            RandomTumblePillingTest_ViewModel result = _Service.EncodeAmend(new RandomTumblePillingTest_ViewModel()
            {
                Main = new RandomTumblePillingTest_Main()
                {
                    ReportNo = ReportNo,
                    Status= "Confirmed",
                    Result = Result
                },

            }, this.UserID);

            RandomTumblePillingTest_ViewModel model = _Service.GetData(new RandomTumblePillingTest_Request() { ReportNo = ReportNo });

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
            RandomTumblePillingTest_ViewModel result = _Service.EncodeAmend(new RandomTumblePillingTest_ViewModel()
            {
                Main = new RandomTumblePillingTest_Main()
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

            RandomTumblePillingTest_ViewModel result = _Service.GetReport(ReportNo, false);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = "/TMP/" + result.TempFileName;

            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ToPDF(string ReportNo)
        {
            RandomTumblePillingTest_ViewModel result = _Service.GetReport(ReportNo, true);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = "/TMP/" + result.TempFileName;

            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        //public JsonResult SendMail(string ReportNo)
        //{
        //    this.CheckSession();
        //    RandomTumblePillingTest_ViewModel result = _Service.GetReport(ReportNo, true);

        //    if (!result.Result)
        //    {
        //        result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
        //        return Json(new { result.Result, ErrMsg = result.ErrorMessage });
        //    }

        //    string reportPath = "/TMP/" + result.TempFileName;

        //    return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, FileName = result.TempFileName });
        //}

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = _Service.SendMail(ReportNo, TO, CC, Subject, Body, Files);
            return Json(result);
        }
    }
}