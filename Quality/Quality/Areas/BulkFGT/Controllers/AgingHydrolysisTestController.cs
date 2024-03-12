using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
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
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class AgingHydrolysisTestController : BaseController
    {
        private AgingHydrolysisTestService _service;
        public AgingHydrolysisTestController()
        {
            _service = new AgingHydrolysisTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.AgingHydrolysisTest,,";
        }


        // GET: BulkFGT/AgingHydrolysisTest
        public ActionResult Index()
        {
            AgingHydrolysisTest_ViewModel model = _service.GetDefaultModel();

            if (TempData["NewSaveModel"] != null)
            {
                model = (AgingHydrolysisTest_ViewModel)TempData["NewSaveModel"];
            }
            else if (TempData["EditSaveModel"] != null)
            {
                model = (AgingHydrolysisTest_ViewModel)TempData["EditSaveModel"];
            }

            return View(model);
        }

        public ActionResult IndexBack(string BrandID, string SeasonID, string StyleID, string Article,string OrderID)
        {
            AgingHydrolysisTest_Request request = new AgingHydrolysisTest_Request()
            {
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
                OrderID = OrderID,
            };

            AgingHydrolysisTest_ViewModel model = _service.GetMainPage(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(AgingHydrolysisTest_Request request)
        {
            AgingHydrolysisTest_ViewModel model = _service.GetMainPage(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult New()
        {
            AgingHydrolysisTest_ViewModel model = _service.GetDefaultModel();

            model.MainData.EditType = "New";
            model.MainData.TimeUnit = "Hour";

            return View("Index", model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(AgingHydrolysisTest_ViewModel requestModel)
        {
            AgingHydrolysisTest_ViewModel model = _service.SaveMainPage(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSaveModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(AgingHydrolysisTest_ViewModel requestModel)
        {
            AgingHydrolysisTest_ViewModel model = _service.SaveMainPage(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSaveModel"] = model;

            return RedirectToAction("Index");
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(AgingHydrolysisTest_ViewModel requestModel)
        {
            AgingHydrolysisTest_ViewModel model = _service.DeleteMainPage(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return RedirectToAction("Index");
        }


        public ActionResult AddDetailRow(int lastNo)
        {
            AgingHydrolysisTest_ViewModel model = new AgingHydrolysisTest_ViewModel();


            List<string> materialTypeOption = new List<string>();
            foreach (var item in model.MaterialType_Source)
            {
                string selected = string.Empty;
                materialTypeOption.Add($"<option {selected} value='{item.Value}'>{item.Text}</option>");
            }

            string oddEven = lastNo % 2 == 0 ? "odd" : "even";
            string html = string.Empty;
            html += $@"
<!--#region Row {lastNo}-->

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input type=""hidden"" class=""detailRowIdx"" name=""name"" value=""{lastNo}"" readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <select class=""CanEdit DetailInput"" id=""DetailList{lastNo}__MaterialType"" name=""DetailList[{lastNo}].MaterialType"">
    {string.Join(Environment.NewLine + "            ", materialTypeOption)}
    </select>
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <!--Received Date-->
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <!--Report Date-->
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <!--Last Update Date-->
</div>
<div class=""DetailDataAreaItem2 colBody Row{lastNo}"">
    <img class='detailDelete' src=""/Image/Icon/Delete.png"" width=""30"">
</div>
<!--#endregion-->
";


            return Content(html);
        }

        public ActionResult OrderIDCheck(string orderID)
        {
            AgingHydrolysisTest_ViewModel model = _service.GetOrderInfo(orderID);

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
                    if (item.Value == model.MainData.Article)
                    {
                        selected = "selected";
                    }
                    Article_Source.Add($"<option {selected} value='{item.Value}'>{item.Text}</option>");
                }
                return Json(new { ErrMsg = string.Empty, Result = model.Result, MainData = model.MainData, ArticleSource = Article_Source });
            }
        }



        public ActionResult Detail(string ReportNo)
        {
            AgingHydrolysisTest_Detail_ViewModel model = _service.GetDetailPage(new AgingHydrolysisTest_Request() { ReportNo = ReportNo });

            ViewBag.FactoryID = this.FactoryID;
            return View(model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "DetailSave")]
        public ActionResult DetailSave(AgingHydrolysisTest_Detail_ViewModel requestModel)
        {
            AgingHydrolysisTest_Detail_ViewModel model = _service.SaveDetailPage(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }


            return RedirectToAction("Detail", new { ReportNo = requestModel.MainDetailData.ReportNo });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode(string ReportNo, string Result)
        {
            CheckSession();
            AgingHydrolysisTest_ViewModel result = _service.EncodeAmend(new AgingHydrolysisTest_Detail()
            {
                ReportNo = ReportNo,
                Status = "Confirmed",
                Result = Result
            }, this.UserID);
            AgingHydrolysisTest_Detail_ViewModel model = _service.GetDetailPage(new AgingHydrolysisTest_Request() { ReportNo = ReportNo });

            return Json(new { result.Result, ErrMsg = result.ErrorMessage, AgingHydrolysisResult = model.MainDetailData.Result });
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Amend(string ReportNo, string Result)
        {
            CheckSession();
            AgingHydrolysisTest_ViewModel result = _service.EncodeAmend(new AgingHydrolysisTest_Detail()
            {
                ReportNo = ReportNo,
                Status = "New",
                Result = string.Empty
            }, this.UserID);

            return Json(new { result.Result, ErrMsg = result.ErrorMessage });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ToExcel(string ReportNo)
        {
            this.CheckSession();

            AgingHydrolysisTest_Detail_ViewModel result = _service.GetReport(ReportNo, false);

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
            this.CheckSession();

            AgingHydrolysisTest_Detail_ViewModel result = _service.GetReport(ReportNo, true);

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
        public JsonResult FailMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = _service.FailSendMail(ReportNo, TO, CC, Subject, Body, Files);
            return Json(result);
        }

    }
}