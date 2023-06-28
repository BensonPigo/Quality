using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using DatabaseObject;
using DatabaseObject.ResultModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class FabricOvenTestController : BaseController
    {
        private IFabricOvenTestService _FabricOvenTestService;
        private List<string> Results = new List<string>() { "Pass", "Fail" };
        private List<string> Temperatures = new List<string>() { "70", "90" };
        private List<string> Times = new List<string>() { "4", "24", "48" };
        public FabricOvenTestController()
        {
            _FabricOvenTestService = new FabricOvenTestService();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.FabricOvenTest,,";
        }

        // GET: BulkFGT/FabricOvenTest
        public ActionResult Index()
        {
            FabricOvenTest_Result model = new FabricOvenTest_Result()
            {
                Main = new FabricOvenTest_Main(),
                Details = new List<FabricOvenTest_Detail>()
            };

            return View(model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult Index(string POID)
        {
            // 21051739BB
            FabricOvenTest_Result model = _FabricOvenTestService.GetFabricOvenTest_Result(POID);
            if (!model.Result)
            {
                model = new FabricOvenTest_Result()
                        {
                            Main = new FabricOvenTest_Main(),
                            Details = new List<FabricOvenTest_Detail>(),
                            ErrorMessage = $@"msg.WithInfo('{ (string.IsNullOrEmpty(model.ErrorMessage) ? string.Empty : model.ErrorMessage.Replace("'", string.Empty))  }');" ,
                        };
            }
            ViewBag.POID = POID;
            UpdateModel(model);
            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        [SessionAuthorizeAttribute]
        public ActionResult IndexBack(string POID)
        {
            FabricOvenTest_Result model = _FabricOvenTestService.GetFabricOvenTest_Result(POID);
            ViewBag.POID = POID;
            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        public ActionResult Detail(string POID, string TestNo,string EditMode, string BrandID)
        {
            FabricOvenTest_Detail_Result model = _FabricOvenTestService.GetFabricOvenTest_Detail_Result(POID, TestNo, BrandID);

            List<SelectListItem> ScaleIDList = new SetListItem().ItemListBinding(model.ScaleIDs);
            List<SelectListItem> ResultChangeList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> ResultStainList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> TimeList = new SetListItem().ItemListBinding(Times);

            if (Convert.ToBoolean(EditMode) && string.IsNullOrEmpty(TestNo))
            {
                model.Main.Status = "New";
            }

            if (TempData["ModelFabricOvenTest"] != null)
            {
                FabricOvenTest_Detail_Result saveResult = (FabricOvenTest_Detail_Result)TempData["ModelFabricOvenTest"];
                model.Main.InspDate = saveResult.Main.InspDate;
                model.Main.Article = saveResult.Main.Article;
                model.Main.Inspector = saveResult.Main.Inspector;
                model.Main.InspectorName = saveResult.Main.InspectorName;
                model.Main.Remark = saveResult.Main.Remark;
                model.Details = saveResult.Details;
                model.Result = saveResult.Result;
                model.ErrorMessage = $@"msg.WithInfo('{  (string.IsNullOrEmpty(saveResult.ErrorMessage) ? string.Empty : saveResult.ErrorMessage.Replace("'", string.Empty))  }');EditMode=true;";
                EditMode = "True";
            }

            ViewBag.ChangeScaleList = ScaleIDList;
            ViewBag.ResultChangeList = ResultChangeList;
            ViewBag.StainingScaleList = ScaleIDList;
            ViewBag.ResultStainList = ResultStainList;
            ViewBag.TemperatureList = TemperatureList;
            ViewBag.TimeList = TimeList;
            ViewBag.EditMode = EditMode;
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.ErrorMessage = string.Empty;
            return View(model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult DetailSave(FabricOvenTest_Detail_Result req)
        {
            req.MDivisionID = this.MDivisionID;

            req.Main.TestBeforePicture = req.Main.TestBeforePicture == null ? null : ImageHelper.ImageCompress(req.Main.TestBeforePicture);
            req.Main.TestAfterPicture = req.Main.TestAfterPicture == null ? null : ImageHelper.ImageCompress(req.Main.TestAfterPicture);

            BaseResult result = _FabricOvenTestService.SaveFabricOvenTestDetail(req, this.UserID);
            if (result.Result)
            {
                // 找地方寫入 TestNo
                req.Main.TestNo = string.IsNullOrEmpty(result.ErrorMessage) ? req.Main.TestNo : result.ErrorMessage;
                return RedirectToAction("Detail", new { POID = req.Main.POID, TestNo = req.Main.TestNo, EditMode = false });
            }

            req.Result = result.Result;
            req.ErrorMessage = result.ErrorMessage;
            TempData["ModelFabricOvenTest"] = req;
            return RedirectToAction("Detail", new { POID = req.Main.POID, TestNo = req.Main.TestNo, EditMode = false });
        }

        [SessionAuthorizeAttribute]
        public JsonResult MainDetailDelete(string ID, string No)
        {
            BaseResult result = _FabricOvenTestService.DeleteOven(ID, No);

            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult SaveMaster(FabricOvenTest_Main Main)
        {
            var result = _FabricOvenTestService.SaveFabricOvenTestMain(Main);       

            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult AddDetailRow(string POID, int lastNO, string GroupNO, string BrandID)
        {
            FabricOvenTest_Detail_Result model = _FabricOvenTestService.GetFabricOvenTest_Detail_Result(POID, "", "");


            List<SelectListItem> ResultChangeList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> ResultStainList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> TimeList = new SetListItem().ItemListBinding(Times);

            int i = lastNO;
            FabricOvenTest_Detail_Detail detail = new FabricOvenTest_Detail_Detail();
            string html = "";
            html += "<tr idx='" + i + "'>";
            html += "<td><input id='Seq" + i + "' idx='" + i + "' type ='hidden'></input> <input id='Details_" + i + "__SubmitDate' name='Details[" + i + "].SubmitDate' class='form-control date-picker width9vw' type='text' value=''></td>";
            html += "<td><input id='Details_" + i + "__OvenGroup' name='Details[" + i + "].OvenGroup' type='number' max='99' maxlength='2' min='0' step='1' oninput='value=OvenGroupCheck(value)' value='" + GroupNO + "'></td>"; // group
            html += "<td style='width: 11vw;'><div style='width:10vw;'><input id='Details_" + i + "__SEQ' name='Details[" + i + "].SEQ' idv='" + i.ToString() + "' class ='InputDetailSEQSelectItem' type='text'  style = 'width: 6vw'> <input id='btnDetailSEQSelectItem'  idv='" + i.ToString() + "' type='button' class='btnDetailSEQSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // seq
            html += "<td style='width: 11vw;'><div style='width:10vw;'><input id='Details_" + i + "__Roll' name='Details[" + i + "].Roll' idv='" + i.ToString() + "' class ='InputDetailRollSelectItem' type='text' style = 'width: 6vw'> <input id='btnDetailRollSelectItem' idv='" + i.ToString() + "' type='button' class='btnDetailRollSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // roll
            html += "<td><input id='Details_" + i + "__Dyelot' name='Details[" + i + "].Dyelot' type='text' readonly='readonly'></td>"; // dyelot
            html += "<td><input id='Details_" + i + "__Refno' name='Details[" + i + "].Refno' type='text' readonly='readonly'></td>"; // Refno
            html += "<td><input id='Details_" + i + "__SCIRefno' name='Details[" + i + "].SCIRefno' type='text' readonly='readonly'></td>"; // SCIRefno
            html += "<td><input id='Details_" + i + "__ColorID' name='Details[" + i + "].ColorID' type='text' readonly='readonly'></td>"; // ColorID
            html += "<td><input  readonly='readonly'  id='Details_" + i + "__Result' name='Details[" + i + "].Result'  class='blue width6vw' type='text'></td>"; // Result

            html += "<td><select id='Details_" + i + "__ChangeScale' name='Details[" + i + "].ChangeScale'>"; // ChangeScale
            foreach (string val in model.ScaleIDs)
            {
                if (val == "4-5")
                {
                    html += "<option selected value='" + val + "'>" + val + "</option>";
                }
                else
                {
                    html += "<option value='" + val + "'>" + val + "</option>";
                }
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultChange' name='Details[" + i + "].ResultChange' >"; // ResultChange
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Details_" + i + "__StainingScale' name='Details[" + i + "].StainingScale'>"; // StainingScale
            foreach (string val in model.ScaleIDs)
            {
                if (val == "4-5")
                {
                    html += "<option selected value='" + val + "'>" + val + "</option>";
                }
                else
                {
                    html += "<option value='" + val + "'>" + val + "</option>";
                }
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultStain' name='Details[" + i + "].ResultStain' >"; // ResultStain
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><input id='Details_" + i + "__Remark' name='Details[" + i + "].Remark' type='text'></td>"; // remark

            html += "<td></td>"; // LastUpdate

            // Temperature
            if (BrandID == "ADIDAS" || BrandID == "REEBOK")
            {
                html += "<td><select id='Details_" + i + "__Temperature' name='Details[" + i + "].Temperature' >"; // Temperature
                foreach (string val in Temperatures)
                {
                    html += "<option value='" + val + "'>" + val + "</option>";
                }
                html += "</select></td>";
            }
            else
            {
                html += "<td>"; 
                html += $@"<input id=""Details_{i}__Temperature"" name=""Details[{i}].Temperature"" onchange=""value=IntCheck(value)"" type=""number"">";                
                html += "</td>";
            }


            html += "<td><select id='Details_" + i + "__Time' name='Details[" + i + "].Time' >"; // Time
            foreach (string val in Times)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><img  class='detailDelete' src='/Image/Icon/Delete.png' width='30'></td>";
            html += "</tr>";

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode_Detail(string POID, string TestNo)
        {
            string ovenTestResult = string.Empty;
            BaseResult result = _FabricOvenTestService.EncodeFabricOvenTestDetail(POID, TestNo,out ovenTestResult);
            return Json(new { result.Result, result.ErrorMessage, ovenTestResult });  
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult FailMail(string ID, string No, string TO, string CC)
        {
            SendMail_Result result = _FabricOvenTestService.SendFailResultMail(TO, CC, ID, No, false);
            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Amend_Detail(string POID, string TestNo)
        {
            BaseResult result = _FabricOvenTestService.AmendFabricOvenTestDetail(POID, TestNo);
            return Json(new { result.Result, ErrorMessage = string.IsNullOrEmpty(result.ErrorMessage) ? string.Empty : result.ErrorMessage.Replace("'", string.Empty) });
        }


        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Report(string ID, string No, bool IsToPDF)
        {
            BaseResult result;
            string FileName;
            if (IsToPDF)
            {
                result = _FabricOvenTestService.ToPdfFabricOvenTestDetail(ID, No, out FileName, false);
            }
            else
            {
                result = _FabricOvenTestService.ToExcelFabricOvenTestDetail(ID, No, out FileName, false);
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult SendMail(string ID, string No)
        {
            this.CheckSession();

            BaseResult result = null;
            string FileName = string.Empty;

            result = _FabricOvenTestService.ToPdfFabricOvenTestDetail(ID, No, out FileName, false);

            if (!result.Result)
            {
                result.ErrorMessage = result.ErrorMessage.ToString();
            }
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;

            return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, reportPath = reportPath, FileName = FileName });
        }
    }
}