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
    public class SalivaFastnessTestController : BaseController
    {
        private SalivaFastnessTestService _Service;
        public SalivaFastnessTestController()
        {
            _Service = new SalivaFastnessTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.SalivaFastnessTest,,";
        }
        // GET: BulkFGT/SalivaFastnessTest
        public ActionResult Index()
        {
            SalivaFastnessTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["NewSaveSalivaFastnessTestModel"] != null)
            {
                model = (SalivaFastnessTest_ViewModel)TempData["NewSaveSalivaFastnessTestModel"];
            }
            else if (TempData["EditSaveSalivaFastnessTestModel"] != null)
            {
                model = (SalivaFastnessTest_ViewModel)TempData["EditSaveSalivaFastnessTestModel"];
            }
            else if (TempData["DeleteSalivaFastnessTestModel"] != null)
            {
                model = (SalivaFastnessTest_ViewModel)TempData["DeleteSalivaFastnessTestModel"];
            }

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            SalivaFastnessTest_Request request = new SalivaFastnessTest_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            SalivaFastnessTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(SalivaFastnessTest_Request request)
        {
            SalivaFastnessTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult New()
        {
            SalivaFastnessTest_ViewModel model = _Service.GetDefaultModel(true);

            model.Main.EditType = "New";

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(SalivaFastnessTest_ViewModel requestModel)
        {
            SalivaFastnessTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSaveSalivaFastnessTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(SalivaFastnessTest_ViewModel requestModel)
        {
            SalivaFastnessTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSaveSalivaFastnessTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(SalivaFastnessTest_ViewModel requestModel)
        {
            SalivaFastnessTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }
            TempData["DeleteSalivaFastnessTestModel"] = model;
            return RedirectToAction("Index");
        }

        public ActionResult AddDetailRow(int lastNo)
        {
            SalivaFastnessTest_ViewModel model = new SalivaFastnessTest_ViewModel();
            model = _Service.GetDefaultModel();

            List<string> ScaleOption = new List<string>();
            foreach (var item in model.Scale_Source)
            {
                string selected = string.Empty;
                if (item.Value == "4-5")
                {
                    selected = "selected";
                }
                ScaleOption.Add($"<option {selected} value='{item.Value}'>{item.Text}</option>");
            }

            string html = string.Empty;
            html += $@"
<!--#region Row {lastNo}-->

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input type=""hidden"" class=""detailRowIdx"" name=""name"" value=""{lastNo}"" readonly=""readonly"">
    <input class="""" id=""DetailList_{lastNo}__EvaluationItem"" name=""DetailList[{lastNo}].EvaluationItem"" type=""text"" value=""Staining"" readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__AllResult"" name=""DetailList[{lastNo}].AllResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <select class=""CanEdit DetailInput"" id=""DetailList_{lastNo}__AcetateScale"" name=""DetailList[{lastNo}].AcetateScale""  onchange=""ScaleChange({lastNo}, this, 'AcetateResult', '4-5')"">
    {string.Join(Environment.NewLine + "            ", ScaleOption)}
    </select>
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <select class=""CanEdit DetailInput"" id=""DetailList_{lastNo}__CottonScale"" name=""DetailList[{lastNo}].CottonScale""  onchange=""ScaleChange({lastNo}, this, 'CottonResult', '4-5')"">
    {string.Join(Environment.NewLine + "            ", ScaleOption)}
    </select>
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <select class=""CanEdit DetailInput"" id=""DetailList_{lastNo}__NylonScale"" name=""DetailList[{lastNo}].NylonScale""  onchange=""ScaleChange({lastNo}, this, 'NylonResult', '4-5')"">
    {string.Join(Environment.NewLine + "            ", ScaleOption)}
    </select>
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <select class=""CanEdit DetailInput"" id=""DetailList_{lastNo}__PolyesterScale"" name=""DetailList[{lastNo}].PolyesterScale""  onchange=""ScaleChange({lastNo}, this, 'PolyesterResult', '4-5')"">
    {string.Join(Environment.NewLine + "            ", ScaleOption)}
    </select>
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <select class=""CanEdit DetailInput"" id=""DetailList_{lastNo}__AcrylicScale"" name=""DetailList[{lastNo}].AcrylicScale""  onchange=""ScaleChange({lastNo}, this, 'AcrylicResult', '4-5')"">
    {string.Join(Environment.NewLine + "            ", ScaleOption)}
    </select>
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <select class=""CanEdit DetailInput"" id=""DetailList_{lastNo}__WoolScale"" name=""DetailList[{lastNo}].WoolScale""  onchange=""ScaleChange({lastNo}, this, 'WoolResult', '4-5')"">
    {string.Join(Environment.NewLine + "            ", ScaleOption)}
    </select>
</div>


<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__AcetateResult"" name=""DetailList[{lastNo}].AcetateResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__CottonResult"" name=""DetailList[{lastNo}].CottonResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__NylonResult"" name=""DetailList[{lastNo}].NylonResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__PolyesterResult"" name=""DetailList[{lastNo}].PolyesterResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__AcrylicResult"" name=""DetailList[{lastNo}].AcrylicResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__WoolResult"" name=""DetailList[{lastNo}].WoolResult"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
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
<!--#endregion-->
";


            return Content(html);
        }

        public ActionResult OrderIDCheck(string orderID)
        {
            SalivaFastnessTest_ViewModel model = _Service.GetOrderInfo(orderID);

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
            SalivaFastnessTest_ViewModel result = _Service.EncodeAmend(new SalivaFastnessTest_ViewModel()
            {
                Main = new SalivaFastnessTest_Main()
                {
                    ReportNo = ReportNo,
                    Status= "Confirmed",
                    Result = Result
                },

            }, this.UserID);

            SalivaFastnessTest_ViewModel model = _Service.GetData(new SalivaFastnessTest_Request() { ReportNo = ReportNo });

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
            SalivaFastnessTest_ViewModel result = _Service.EncodeAmend(new SalivaFastnessTest_ViewModel()
            {
                Main = new SalivaFastnessTest_Main()
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

            SalivaFastnessTest_ViewModel result = _Service.GetReport(ReportNo, false);

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
            SalivaFastnessTest_ViewModel result = _Service.GetReport(ReportNo, true);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + result.TempFileName;

            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        public JsonResult SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = _Service.SendMail(ReportNo, TO, CC, Subject, Body, Files);
            return Json(result);
        }
    }
}