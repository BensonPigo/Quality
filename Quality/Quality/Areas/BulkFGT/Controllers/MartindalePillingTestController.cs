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
    public class MartindalePillingTestController : BaseController
    {
        private MartindalePillingTestService _Service;
        public MartindalePillingTestController()
        {
            _Service = new MartindalePillingTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.MartindalePillingTest,,";
        }
        // GET: BulkFGT/MartindalePillingTest
        public ActionResult Index()
        {
            MartindalePillingTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["NewSaveMartindalePillingTestModel"] != null)
            {
                model = (MartindalePillingTest_ViewModel)TempData["NewSaveMartindalePillingTestModel"];
            }
            else if (TempData["EditSaveMartindalePillingTestModel"] != null)
            {
                model = (MartindalePillingTest_ViewModel)TempData["EditSaveMartindalePillingTestModel"];
            }
            else if (TempData["DeleteMartindalePillingTestModel"] != null)
            {
                model = (MartindalePillingTest_ViewModel)TempData["DeleteMartindalePillingTestModel"];
            }

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            MartindalePillingTest_Request request = new MartindalePillingTest_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            MartindalePillingTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(MartindalePillingTest_Request request)
        {
            MartindalePillingTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult New()
        {
            MartindalePillingTest_ViewModel model = _Service.GetDefaultModel(true);

            model.Main.EditType = "New";

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(MartindalePillingTest_ViewModel requestModel)
        {
            MartindalePillingTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSaveMartindalePillingTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(MartindalePillingTest_ViewModel requestModel)
        {
            MartindalePillingTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSaveMartindalePillingTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(MartindalePillingTest_ViewModel requestModel)
        {
            MartindalePillingTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }
            TempData["DeleteMartindalePillingTestModel"] = model;
            return RedirectToAction("Index");
        }

        public ActionResult AddDetailRow(int lastNo)
        {
            MartindalePillingTest_ViewModel model = new MartindalePillingTest_ViewModel();
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
            MartindalePillingTest_ViewModel model = _Service.GetOrderInfo(orderID);

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

                List<string> TestStandard_Source = new List<string>();
                if (model.Main.FabricType.ToUpper() == "WOVEN")
                {
                    TestStandard_Source.Add($"<option selected value='W1'>W1</option>");
                    TestStandard_Source.Add($"<option value='W2'>W2</option>");
                    TestStandard_Source.Add($"<option value='W3'>W3</option>");
                }
                if (model.Main.FabricType.ToUpper() == "KNIT")
                {
                    TestStandard_Source.Add($"<option selected value='K1'>K1</option>");
                    TestStandard_Source.Add($"<option value='K2'>K2</option>");
                    TestStandard_Source.Add($"<option value='K3'>K3</option>");
                }

                return Json(new { ErrMsg = string.Empty, Result = model.Result, Main = model.Main, ArticleSource = Article_Source , TestStandardSource = TestStandard_Source });
            }
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode(string ReportNo, string Result)
        {
            CheckSession();
            MartindalePillingTest_ViewModel result = _Service.EncodeAmend(new MartindalePillingTest_ViewModel()
            {
                Main = new MartindalePillingTest_Main()
                {
                    ReportNo = ReportNo,
                    Status= "Confirmed",
                    Result = Result
                },

            }, this.UserID);

            MartindalePillingTest_ViewModel model = _Service.GetData(new MartindalePillingTest_Request() { ReportNo = ReportNo });

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
            MartindalePillingTest_ViewModel result = _Service.EncodeAmend(new MartindalePillingTest_ViewModel()
            {
                Main = new MartindalePillingTest_Main()
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

            MartindalePillingTest_ViewModel result = _Service.GetReport(ReportNo, false);

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
            MartindalePillingTest_ViewModel result = _Service.GetReport(ReportNo, true);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + result.TempFileName;

            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        //public JsonResult SendMailToMR(string ReportNo)
        //{
        //    this.CheckSession();
        //    MartindalePillingTest_ViewModel result = _Service.GetReport(ReportNo, true);

        //    if (!result.Result)
        //    {
        //        result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
        //        return Json(new { result.Result, ErrMsg = result.ErrorMessage });
        //    }

        //    string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + result.TempFileName;

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