using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;
using static System.Net.Mime.MediaTypeNames;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class EvaporationRateTestController : BaseController
    {
        private EvaporationRateTestService _Service;
        public EvaporationRateTestController()
        {
            _Service = new EvaporationRateTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.EvaporationRateTest,,";
        }
        // GET: BulkFGT/EvaporationRateTest
        public ActionResult Index()
        {
            EvaporationRateTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["NewSaveEvaporationRateTestModel"] != null)
            {
                model = (EvaporationRateTest_ViewModel)TempData["NewSaveEvaporationRateTestModel"];
            }
            else if (TempData["EditSaveEvaporationRateTestModel"] != null)
            {
                model = (EvaporationRateTest_ViewModel)TempData["EditSaveEvaporationRateTestModel"];
            }

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            EvaporationRateTest_Request request = new EvaporationRateTest_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            EvaporationRateTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(EvaporationRateTest_Request request)
        {
            EvaporationRateTest_ViewModel model = _Service.GetData(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult New()
        {
            EvaporationRateTest_ViewModel model = _Service.GetDefaultModel(true);

            model.Main.EditType = "New";

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(EvaporationRateTest_ViewModel requestModel)
        {
            EvaporationRateTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSaveEvaporationRateTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(EvaporationRateTest_ViewModel requestModel)
        {
            EvaporationRateTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSaveEvaporationRateTestModel"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(EvaporationRateTest_ViewModel requestModel)
        {
            EvaporationRateTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }
            return RedirectToAction("Index");
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        public ActionResult AddTimeRow(int TimeListIdx, string detailType, string specimenID, int lastTime, string parentTabID)
        {
            MockupOven_ViewModel model = new MockupOven_ViewModel();

            string html = "";
            html += $@"
<div class=""divTime10"">
    <input class="""" id=""TimeList_{TimeListIdx}__Ukey"" name = ""TimeList[{TimeListIdx}].Ukey"" style = ""width: 100 % "" type = ""hidden"" value = ""0"" >
    <input class="""" id=""TimeList_{TimeListIdx}__SpecimenUkey"" name = ""TimeList[{TimeListIdx}].SpecimenUkey"" style = ""width:100% "" type = ""hidden"" value = ""0"" >
    <input class="""" id=""TimeList_{TimeListIdx}__SpecimenID"" name=""TimeList[{TimeListIdx}].SpecimenID"" style=""width:100%"" type=""hidden"" value=""{specimenID}"">
    <input class="""" id=""TimeList_{TimeListIdx}__DetailType"" name=""TimeList[{TimeListIdx}].DetailType"" style=""width:100%"" type=""hidden"" value=""{detailType}"">
    
    <input class="""" id=""TimeList_{TimeListIdx}__InitialTime"" name=""TimeList[{TimeListIdx}].InitialTime"" style=""width:100%"" type=""hidden"" value=""0"">
    <input class="""" id=""TimeList_{TimeListIdx}__InitialTimeUkey"" name=""TimeList[{TimeListIdx}].InitialTimeUkey"" style=""width:100%"" type=""hidden"" value=""0"">
    <input class="""" id=""TimeList_{TimeListIdx}__IsInitialMass"" name=""TimeList[{TimeListIdx}].IsInitialMass"" style=""width:100%"" type=""hidden"" value=""False"">
    <input class=""Time"" id=""TimeList_{TimeListIdx}__Time"" name=""TimeList[{TimeListIdx}].Time"" detailtype=""{detailType}"" specimenid=""{specimenID}""  style=""width:100%"" type=""text"" value=""{lastTime}"" readonly>
</div>
<div class=""divTime10"">        
    <input detailtype=""{detailType}"" isinitialmass=""False"" specimenid=""{specimenID}"" time=""{lastTime}"" class=""CanEdit Mass"" id=""TimeList_{TimeListIdx}__Mass"" name=""TimeList[{TimeListIdx}].Mass"" onchange=""value=MassCheck(value);AutoUpdateTime(this)"" style=""width:100%"" type=""text"" value=""0"">
</div>                                                                
<div class=""divTime30"">
    <input detailtype=""{detailType}"" parentTabID=""{parentTabID}"" TimeListIdx=""{TimeListIdx}"" initialtime=""0"" specimenid=""{specimenID}"" time=""{lastTime}"" class=""Evaporation"" id=""TimeList_{TimeListIdx}__Evaporation"" name=""TimeList[{TimeListIdx}].Evaporation"" style=""width:100%"" type=""text"" value=""0"">
</div>
<div class=""divTime40"">
    <input class="""""" id = ""TimeList_{TimeListIdx}__LastUpadate"" name = ""TimeList[{TimeListIdx}].LastUpadate"" style = ""width: 100 % "" type = ""text"" value = """" >    
</div>
<div class=""divTime10"">
    <input type=""hidden"" class=""detailTimeListIdx"" name=""name"" value=""{TimeListIdx}"" />
	<img class=""detailDelete"" src=""/Image/Icon/Delete.png"" width=""30"" style=""min-width:30px"" />
</div>
";

            return Content(html);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        public ActionResult AddRateRow(int RateListIdx, int SpecimenListIdx, string detailType, string specimenID, string lastRateName, int minuendTime, int subtrahendTime, string parentTabID)
        {

            MockupOven_ViewModel model = new MockupOven_ViewModel();
            lastRateName = "R" + lastRateName;
            string html = "";
            html += $"";
            html += $@"
<div class=""divRateInner1"">
    <input class="""" id=""RateList_{RateListIdx}__Ukey"" name = ""RateList[{RateListIdx}].Ukey"" style = ""width: 100 % "" type = ""hidden"" value = ""0"" >
    <input class="""" id=""RateList_{RateListIdx}__SpecimenUkey"" name=""RateList[{RateListIdx}].SpecimenUkey"" style=""width:100%"" type=""hidden"" value=""0"">
    <input class="""" id=""RateList_{RateListIdx}__Subtrahend_Time"" name=""RateList[{RateListIdx}].Subtrahend_Time"" style=""width:100%"" type=""hidden"" value=""{subtrahendTime}"">
    <input class="""" id=""RateList_{RateListIdx}__Minuend_Time"" name=""RateList[{RateListIdx}].Minuend_Time"" style=""width:100%"" type=""hidden"" value=""{minuendTime}"">
    <input class="""" id=""RateList_{RateListIdx}__Subtrahend_TimeUkey"" name=""RateList[{RateListIdx}].Subtrahend_TimeUkey"" style=""width:100%"" type=""hidden"" value=""0"">
    <input class="""" id=""RateList_{RateListIdx}__Minuend_TimeUkey"" name=""RateList[{RateListIdx}].Minuend_TimeUkey"" style=""width:100%"" type=""hidden"" value=""0"">
    <input class="""" id=""RateList_{RateListIdx}__DetailType"" name=""RateList[{RateListIdx}].DetailType"" style=""width:100%"" type=""hidden"" value=""{detailType}"">
    <input class="""" id=""RateList_{RateListIdx}__SpecimenID"" name=""RateList[{RateListIdx}].SpecimenID"" style=""width:100%"" type=""hidden"" value=""{specimenID}"">
    <input class="""" id=""RateList_{RateListIdx}__Ratio"" name=""RateList[{RateListIdx}].Ratio"" style=""width:100%"" type=""hidden"" value=""6"">

    <input class="""" id=""RateList_{RateListIdx}__RateName"" name=""RateList[{RateListIdx}].RateName"" style=""width:100%"" type=""text"" value=""{lastRateName}"">
</div>
<div class=""divRateInner1"">
    <input type=""hidden"" class=""detailRateListIdx"" name=""name"" value=""{RateListIdx}"" />
    <input detailtype = ""{detailType}"" parentTabID=""{parentTabID}"" RateListIdx=""{RateListIdx}"" minuend-time=""{minuendTime}"" SpecimenListIdx=""{SpecimenListIdx}"" ratio=""6"" RateName=""{lastRateName}"" specimenid=""{specimenID}"" subtrahend-time=""{subtrahendTime}"" class=""RateValue"" id=""RateList_{RateListIdx}__Value"" name=""RateList[{RateListIdx}].Value"" style=""width:100%"" type=""text"" value=""0"">
</div>                                                                 
";

            return Content(html);
        }
        public ActionResult OrderIDCheck(string orderID)
        {
            EvaporationRateTest_ViewModel model = _Service.GetOrderInfo(orderID);

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
            EvaporationRateTest_ViewModel result = _Service.EncodeAmend(new EvaporationRateTest_ViewModel()
            {
                Main = new EvaporationRateTest_Main()
                {
                    ReportNo = ReportNo,
                    Status = "Confirmed",
                    Result = Result
                },

            }, this.UserID);

            EvaporationRateTest_ViewModel model = _Service.GetData(new EvaporationRateTest_Request() { ReportNo = ReportNo });

            if (model.Main.Result == "Fail")
            {
                return Json(new { result.Result, ErrMsg = result.ErrorMessage ,Action="FailMail()" });
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
            EvaporationRateTest_ViewModel result = _Service.EncodeAmend(new EvaporationRateTest_ViewModel()
            {
                Main = new EvaporationRateTest_Main()
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

            EvaporationRateTest_ViewModel result = _Service.GetReport(ReportNo, false);

            string filename = result.TempFileName;
            byte[] fileBytes = System.IO.File.ReadAllBytes(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", result.TempFileName));

            // 設置回應為文件下載
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ToPDF(string ReportNo)
        {
            EvaporationRateTest_ViewModel result = _Service.GetReport(ReportNo, true);

            string filename = result.TempFileName;
            byte[] fileBytes = System.IO.File.ReadAllBytes(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", result.TempFileName));

            // 設置回應為文件下載
            return File(fileBytes, "application/pdf", filename);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = _Service.SendMail(ReportNo, TO, CC, Subject, Body, Files);
            return Json(result);
        }

    }
}