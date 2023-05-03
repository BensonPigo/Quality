using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.RequestModel;
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
    public class PhenolicYellowTestController : BaseController
    {
        private PhenolicYellowTestService _Service;
        public PhenolicYellowTestController()
        {
            _Service = new PhenolicYellowTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.PhenolicYellowTest,,";
        }
        // GET: BulkFGT/PhenolicYellowTest
        public ActionResult Index()
        {
            PhenolicYellowTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["NewSavePhenolicYellowTestModel"] != null)
            {
                model = (PhenolicYellowTest_ViewModel)TempData["NewSavePhenolicYellowTestModel"];
            }
            else if (TempData["EditSavePhenolicYellowTestModel"] != null)
            {
                model = (PhenolicYellowTest_ViewModel)TempData["EditSavePhenolicYellowTestModel"];
            }
            else if (TempData["DeletePhenolicYellowTestModel"] != null)
            {
                model = (PhenolicYellowTest_ViewModel)TempData["DeletePhenolicYellowTestModel"];
            }

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            PhenolicYellowTest_Request request = new PhenolicYellowTest_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            PhenolicYellowTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(PhenolicYellowTest_Request request)
        {
            PhenolicYellowTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult New()
        {
            PhenolicYellowTest_ViewModel model = _Service.GetDefaultModel(true);

            model.Main.EditType = "New";

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(PhenolicYellowTest_ViewModel requestModel)
        {
            PhenolicYellowTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSavePhenolicYellowTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(PhenolicYellowTest_ViewModel requestModel)
        {
            PhenolicYellowTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSavePhenolicYellowTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(PhenolicYellowTest_ViewModel requestModel)
        {
            PhenolicYellowTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }

            TempData["DeletePhenolicYellowTestModel"] = model;
            return RedirectToAction("Index");
            //return View("Index", model);
        }

        public ActionResult AddDetailRow(int lastNo)
        {
            PhenolicYellowTest_ViewModel model = new PhenolicYellowTest_ViewModel();
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
    <input class="""" id=""DetailList_{lastNo}__Dyelot"" name=""DetailList[{lastNo}].Dyelot"" type=""text"" value="""" readonly=""readonly"" onclick=""RollDyelotWindow(this)"" placeholder = ""Click"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__Roll"" name=""DetailList[{lastNo}].Roll"" type=""text"" value="""" readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <select class=""CanEdit DetailInput"" id=""DetailList{lastNo}__Scale"" name=""DetailList[{lastNo}].Scale""  onchange=""ScaleChange({lastNo}, this, 'Result', '4-5')"">
    {string.Join(Environment.NewLine + "            ", ScaleOption)}
    </select>
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""DetailList_{lastNo}__Result"" name=""DetailList[{lastNo}].Result"" type=""text"" value=""Pass"" style=""color:blue"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class=""CanEdit"" id=""DetailList_{lastNo}__Remark"" name=""DetailList[{lastNo}].Remark"" type=""text"" value="""">
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
            PhenolicYellowTest_ViewModel model = _Service.GetOrderInfo(orderID);

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
            PhenolicYellowTest_ViewModel result = _Service.EncodeAmend(new PhenolicYellowTest_ViewModel()
            {
                Main = new PhenolicYellowTest_Main()
                {
                    ReportNo = ReportNo,
                    Status= "Confirmed",
                    Result = Result
                },

            }, this.UserID);

            return Json(new { result.Result, ErrMsg = result.ErrorMessage });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Amend(string ReportNo, string Result)
        {
            CheckSession();
            PhenolicYellowTest_ViewModel result = _Service.EncodeAmend(new PhenolicYellowTest_ViewModel()
            {
                Main = new PhenolicYellowTest_Main()
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

            PhenolicYellowTest_ViewModel result = _Service.GetReport(ReportNo, false);

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
            PhenolicYellowTest_ViewModel result = _Service.GetReport(ReportNo, true);

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
            PhenolicYellowTest_ViewModel result = _Service.GetReport(ReportNo, false);

            if (!result.Result)
            {
                result.ErrorMessage = $@"msg.WithInfo(""{result.ErrorMessage.Replace("'", string.Empty)}"");";
                return Json(new { result.Result, ErrMsg = result.ErrorMessage });
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + result.TempFileName;

            return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, FileName = result.TempFileName });
        }
    }
}