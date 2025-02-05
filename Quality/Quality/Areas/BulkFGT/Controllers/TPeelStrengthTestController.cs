using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Microsoft.Office.Interop.Excel;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class TPeelStrengthTestController : BaseController
    {
        private TPeelStrengthTestService _Service;

        public TPeelStrengthTestController()
        {
            _Service = new TPeelStrengthTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.TPeelStrengthTest,,";
        }
        // GET: BulkFGT/TPeelStrengthTest
        public ActionResult Index()
        {
            TPeelStrengthTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["NewSaveTPeelStrengthTestModel"] != null)
            {
                model = (TPeelStrengthTest_ViewModel)TempData["NewSaveTPeelStrengthTestModel"];
            }
            else if (TempData["EditSaveTPeelStrengthTestModel"] != null)
            {
                model = (TPeelStrengthTest_ViewModel)TempData["EditSaveTPeelStrengthTestModel"];
            }
            else if (TempData["DeleteTPeelStrengthTestModel"] != null)
            {
                model = (TPeelStrengthTest_ViewModel)TempData["DeleteTPeelStrengthTestModel"];
            }

            return View(model);
        }


        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            TPeelStrengthTest_Request request = new TPeelStrengthTest_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            TPeelStrengthTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(TPeelStrengthTest_Request request)
        {
            TPeelStrengthTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult New()
        {
            TPeelStrengthTest_ViewModel model = _Service.GetDefaultModel(true);

            model.DetailList = new List<TPeelStrengthTest_Detail>()
                {
                    new TPeelStrengthTest_Detail()
                    {
                        EvaluationItem = "Before",
                        AllResult = "Pass",
                        WarpValue=(decimal)0.5,
                        WarpResult="Pass",
                        WeftValue=(decimal)0.5,
                        WeftResult="Pass"
                    },
                    new TPeelStrengthTest_Detail()
                    {
                        EvaluationItem = "After",
                        AllResult = "Pass",
                        WarpValue=(decimal)0.5,
                        WarpResult="Pass",
                        WeftValue=(decimal)0.5,
                        WeftResult="Pass"
                    }
            };

            model.Main.EditType = "New";

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(TPeelStrengthTest_ViewModel requestModel)
        {
            TPeelStrengthTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSaveTPeelStrengthTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(TPeelStrengthTest_ViewModel requestModel)
        {
            TPeelStrengthTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSaveTPeelStrengthTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(TPeelStrengthTest_ViewModel requestModel)
        {
            TPeelStrengthTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }
            TempData["DeleteTPeelStrengthTestModel"] = model;
            return RedirectToAction("Index");
        }

        public ActionResult AddDetailRow(int lastNo)
        {
            TPeelStrengthTest_ViewModel model = new TPeelStrengthTest_ViewModel();
            model = _Service.GetDefaultModel();

            int lastNoNext = lastNo + 1;
            string html = string.Empty;
            html += $@"
<!--#region Row {lastNo}-->

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input type=""hidden"" class=""detailRowIdx"" name=""name"" value=""{lastNo}"" readonly=""readonly"">
    <input class="""" id=""DetailList_{lastNo}__EvaluationItem"" name=""DetailList[{lastNo}].EvaluationItem"" type=""text"" value=""Before"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__AllResult"" name=""DetailList[{lastNo}].AllResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class=""CanEdit"" id=""DetailList_{lastNo}__WarpValue"" name=""DetailList[{lastNo}].WarpValue"" type=""number"" value=""0.5"" step=""0.01"" onchange = ""value=ValueCheck(value,'{lastNo}','WarpResult')"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__WarpResult"" name=""DetailList[{lastNo}].WarpResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class=""CanEdit"" id=""DetailList_{lastNo}__WeftValue"" name=""DetailList[{lastNo}].WeftValue"" type=""number"" value=""0.5""  step=""0.01"" onchange = ""value=ValueCheck(value,'{lastNo}','WeftResult')"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__WeftResult"" name=""DetailList[{lastNo}].WeftResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem2 colBody Row{lastNo}"">
    <input class=""CanEdit"" id=""DetailList_{lastNo}__Remark"" name=""DetailList[{lastNo}].Remark"" type=""text"" value="""">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <!--Last Update Date-->
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <img class='detailDelete' src=""/Image/Icon/Delete.png"" width=""30"">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNoNext}"">
    <input type=""hidden"" class=""detailRowIdx"" name=""name"" value=""{lastNoNext}"" readonly=""readonly"">
    <input class="""" id=""DetailList_{lastNoNext}__EvaluationItem"" name=""DetailList[{lastNoNext}].EvaluationItem"" type=""text"" value=""After"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNoNext}"">
    <input class="""" id=""DetailList_{lastNoNext}__AllResult"" name=""DetailList[{lastNoNext}].AllResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNoNext}"">
    <input class=""CanEdit"" id=""DetailList_{lastNoNext}__WarpValue"" name=""DetailList[{lastNoNext}].WarpValue"" type=""number"" value=""0.5"" step=""0.01"" onchange = ""value=ValueCheck(value,'{lastNoNext}','WarpResult')"" >
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNoNext}"">
    <input class="""" id=""DetailList_{lastNoNext}__WarpResult"" name=""DetailList[{lastNoNext}].WarpResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNoNext}"">
    <input class=""CanEdit"" id=""DetailList_{lastNoNext}__WeftValue"" name=""DetailList[{lastNoNext}].WeftValue"" type=""number"" value=""0.5"" step=""0.01"" onchange = ""value=ValueCheck(value,'{lastNoNext}','WeftResult')"" >
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNoNext}"">
    <input class="""" id=""DetailList_{lastNoNext}__WeftResult"" name=""DetailList[{lastNoNext}].WeftResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem2 colBody Row{lastNoNext}"">
    <input class=""CanEdit"" id=""DetailList_{lastNoNext}__Remark"" name=""DetailList[{lastNoNext}].Remark"" type=""text"" value="""">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNoNext}"">
    <!--Last Update Date-->
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNoNext}"">
    <img class='detailDelete' src=""/Image/Icon/Delete.png"" width=""30"">
</div>

<!--#endregion-->
";


            return Content(html);
        }

        public ActionResult OrderIDCheck(string orderID)
        {
            TPeelStrengthTest_ViewModel model = _Service.GetOrderInfo(orderID);

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
            TPeelStrengthTest_ViewModel result = _Service.EncodeAmend(new TPeelStrengthTest_ViewModel()
            {
                Main = new TPeelStrengthTest_Main()
                {
                    ReportNo = ReportNo,
                    Status = "Confirmed",
                    Result = Result
                },

            }, this.UserID);

            TPeelStrengthTest_ViewModel model = _Service.GetData(new TPeelStrengthTest_Request() { ReportNo = ReportNo });

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
            TPeelStrengthTest_ViewModel result = _Service.EncodeAmend(new TPeelStrengthTest_ViewModel()
            {
                Main = new TPeelStrengthTest_Main()
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

            TPeelStrengthTest_ViewModel result =  _Service.GetReport(ReportNo, false);

            string filename = result.TempFileName;
            byte[] fileBytes = System.IO.File.ReadAllBytes(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", result.TempFileName));

            // 設置回應為文件下載
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ToPDF(string ReportNo)
        {
            TPeelStrengthTest_ViewModel result = _Service.GetReport(ReportNo, true);

            string filename = result.TempFileName;
            byte[] fileBytes = System.IO.File.ReadAllBytes(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", result.TempFileName));

            // 設置回應為文件下載
            return File(fileBytes, "application/pdf", filename);
        }
        public JsonResult SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = _Service.SendMail(ReportNo, TO, CC, Subject, Body, Files);
            return Json(result);
        }
    }
}