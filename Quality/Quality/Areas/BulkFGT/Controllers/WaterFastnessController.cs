using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject;
using DatabaseObject.ResultModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;


namespace Quality.Areas.BulkFGT.Controllers
{
    public class WaterFastnessController : BaseController
    {
        public WaterFastnessService _WaterFastnessService;
        private List<string> Results = new List<string>() { "Pass", "Fail" };
        private List<string> Temperatures = new List<string>() { "35", "36", "37", "38", "39" };
        private List<string> Times = new List<string>() { "4", "8", "12", "24" };

        public WaterFastnessController()
        {
            _WaterFastnessService = new WaterFastnessService();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.WaterFastness,,";
        }

        // GET: BulkFGT/WaterFastness
        public ActionResult Index()
        {
            WaterFastness_Result model = new WaterFastness_Result()
            {
                Main = new WaterFastness_Main(),
                Details = new List<WaterFastness_Detail>()
            };

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(string POID)
        {
            // 21051739BB
            WaterFastness_Result model = _WaterFastnessService.GetWaterFastness_Result(POID);
            if (!model.Result)
            {
                model = new WaterFastness_Result()
                {
                    Main = new WaterFastness_Main(),
                    Details = new List<WaterFastness_Detail>(),
                    ErrorMessage = $@"msg.WithInfo('{model.ErrorMessage}');",
                };
            }
            ViewBag.POID = POID;
            UpdateModel(model);
            return View(model);
        }
        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexBack(string POID)
        {
            WaterFastness_Result model = _WaterFastnessService.GetWaterFastness_Result(POID);
            ViewBag.POID = POID;
            return View("Index", model);
        }
        public ActionResult Detail(string POID, string TestNo, string EditMode)
        {
            WaterFastness_Detail_Result model = _WaterFastnessService.GetWaterFastness_Detail_Result(POID, TestNo);

            List<SelectListItem> ScaleIDList = new SetListItem().ItemListBinding(model.ScaleIDs);
            List<SelectListItem> ResultChangeList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> ResultStainList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> TimeList = new SetListItem().ItemListBinding(Times);

            if (Convert.ToBoolean(EditMode) && string.IsNullOrEmpty(TestNo))
            {
                model.Main.Status = "New";
            }

            if (TempData["Model"] != null)
            {
                WaterFastness_Detail_Result saveResult = (WaterFastness_Detail_Result)TempData["Model"];
                model.Main.InspDate = saveResult.Main.InspDate;
                model.Main.Article = saveResult.Main.Article;
                model.Main.Inspector = saveResult.Main.Inspector;
                model.Main.InspectorName = saveResult.Main.InspectorName;
                model.Main.Remark = saveResult.Main.Remark;
                model.Details = saveResult.Details;
                model.Result = saveResult.Result;
                model.ErrorMessage = $@"msg.WithInfo(""{saveResult.ErrorMessage.ToString() }"");EditMode=true;";
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
        public ActionResult DetailSave(WaterFastness_Detail_Result req)
        {
            BaseResult result = _WaterFastnessService.SaveWaterFastnessDetail(req, this.UserID);
            if (result.Result)
            {
                // 找地方寫入 TestNo
                req.Main.TestNo = string.IsNullOrEmpty(result.ErrorMessage) ? req.Main.TestNo : result.ErrorMessage;
                return RedirectToAction("Detail", new { POID = req.Main.POID, TestNo = req.Main.TestNo, EditMode = false });
            }

            req.Result = result.Result;
            req.ErrorMessage = result.ErrorMessage;
            TempData["Model"] = req;
            return RedirectToAction("Detail", new { POID = req.Main.POID, TestNo = req.Main.TestNo, EditMode = false });
        }
        public JsonResult MainDetailDelete(string ID, string No)
        {
            BaseResult result = _WaterFastnessService.DeleteWaterFastness(ID, No);

            return Json(result);
        }
        [HttpPost]
        public JsonResult SaveMaster(WaterFastness_Main Main)
        {
            var result = _WaterFastnessService.SaveWaterFastnessMain(Main);

            return Json(result);
        }

        [HttpPost]
        public ActionResult AddDetailRow(string POID, int lastNO, string GroupNO)
        {
            WaterFastness_Detail_Result model = _WaterFastnessService.GetWaterFastness_Detail_Result(POID, "");


            List<SelectListItem> ResultChangeList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> ResultStainList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> TimeList = new SetListItem().ItemListBinding(Times);

            int i = lastNO;
            WaterFastness_Detail_Detail detail = new WaterFastness_Detail_Detail();
            string html = "";
            html += "<tr idx='" + i + "'>";
            html += "<td><input id='Seq" + i + "' idx='" + i + "' type ='hidden'></input> <input id='Details_" + i + "__SubmitDate' name='Details[" + i + "].SubmitDate' class='form-control date-picker' type='text' value=''></td>";
            html += "<td><input id='Details_" + i + "__WaterFastnessGroup' name='Details[" + i + "].WaterFastnessGroup' type='number' max='99' maxlength='2' min='0' step='1' oninput='value=WaterFastnessGroupCheck(value)' value='" + GroupNO + "'></td>"; // group
            html += "<td style='width: 11vw;'><div style='width:10vw;'><input id='Details_" + i + "__SEQ' name='Details[" + i + "].SEQ' idv='" + i.ToString() + "' class ='InputDetailSEQSelectItem' type='text'  style = 'width: 6vw'> <input id='btnDetailSEQSelectItem'  idv='" + i.ToString() + "' type='button' class='btnDetailSEQSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // seq
            html += "<td style='width: 11vw;'><div style='width:10vw;'><input id='Details_" + i + "__Roll' name='Details[" + i + "].Roll' idv='" + i.ToString() + "' class ='InputDetailRollSelectItem' type='text' style = 'width: 6vw'> <input id='btnDetailRollSelectItem' idv='" + i.ToString() + "' type='button' class='btnDetailRollSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // roll
            html += "<td><input id='Details_" + i + "__Dyelot' name='Details[" + i + "].Dyelot' type='text' readonly='readonly'></td>"; // dyelot
            html += "<td><input id='Details_" + i + "__Refno' name='Details[" + i + "].Refno' type='text' readonly='readonly'></td>"; // Refno
            html += "<td><input id='Details_" + i + "__SCIRefno' name='Details[" + i + "].SCIRefno' type='text' readonly='readonly'></td>"; // SCIRefno
            html += "<td><input id='Details_" + i + "__ColorID' name='Details[" + i + "].ColorID' type='text' readonly='readonly'></td>"; // ColorID
            html += "<td><input  readonly='readonly'  id='Details_" + i + "__Result' name='Details[" + i + "].Result'  class='detailResultColor' type='text'></td>"; // Result

            html += "<td><select id='Details_" + i + "__ChangeScale' name='Details[" + i + "].ChangeScale'>"; // ChangeScale
            foreach (string val in model.ScaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultChange' name='Details[" + i + "].ResultChange' >"; // ResultAcetate
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Details_" + i + "__AcetateScale' name='Details[" + i + "].AcetateScale'>"; // AcetateScale
            foreach (string val in model.ScaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultAcetate' name='Details[" + i + "].ResultAcetate' >"; // ResultAcetate
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Details_" + i + "__CottonScale' name='Details[" + i + "].CottonScale'>"; // CottonScale
            foreach (string val in model.ScaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultCotton' name='Details[" + i + "].ResultCotton' >"; // ResultCotton
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Details_" + i + "__NylonScale' name='Details[" + i + "].NylonScale'>"; // NylonScale
            foreach (string val in model.ScaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultNylon' name='Details[" + i + "].ResultNylon' >"; // ResultNylon
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Details_" + i + "__PolyesterScale' name='Details[" + i + "].PolyesterScale'>"; // PolyesterScale
            foreach (string val in model.ScaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultPolyester' name='Details[" + i + "].ResultPolyester' >"; // ResultPolyester
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Details_" + i + "__AcrylicScale' name='Details[" + i + "].AcrylicScale'>"; // AcrylicScale
            foreach (string val in model.ScaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultAcrylic' name='Details[" + i + "].ResultAcrylic' >"; // ResultAcrylic
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Details_" + i + "__WoolScale' name='Details[" + i + "].WoolScale'>"; // WoolScale
            foreach (string val in model.ScaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultWool' name='Details[" + i + "].ResultWool' >"; // ResultWool
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";


            html += "<td><input id='Details_" + i + "__Remark' name='Details[" + i + "].Remark' type='text'></td>"; // remark

            html += "<td></td>"; // LastUpdate


            //html += "<td><select id='Details_" + i + "__Temperature' name='Details[" + i + "].Temperature' >"; // Temperature
            //foreach (string val in Temperatures)
            //{
            //    html += "<option value='" + val + "'>" + val + "</option>";
            //}
            //html += "</select></td>";


            //html += "<td><select id='Details_" + i + "__Time' name='Details[" + i + "].Time' >"; // Time
            //foreach (string val in Times)
            //{
            //    html += "<option value='" + val + "'>" + val + "</option>";
            //}
            //html += "</select></td>";

            html += "<td><img  class='detailDelete' src='/Image/Icon/Delete.png' width='30'></td>";
            html += "</tr>";

            return Content(html);
        }

        [HttpPost]
        public JsonResult Encode_Detail(string POID, string TestNo)
        {
            string waterFastnessResult = string.Empty;
            BaseResult result = _WaterFastnessService.EncodeWaterFastnessDetail(POID, TestNo, out waterFastnessResult);
            return Json(new { result.Result, result.ErrorMessage, waterFastnessResult });
        }

        [HttpPost]
        public JsonResult FailMail(string ID, string No, string TO, string CC)
        {
            SendMail_Result result = _WaterFastnessService.SendFailResultMail(TO, CC, ID, No, false);
            return Json(result);
        }
        [HttpPost]
        public JsonResult Amend_Detail(string POID, string TestNo)
        {
            BaseResult result = _WaterFastnessService.AmendWaterFastnessDetail(POID, TestNo);
            return Json(new { result.Result, result.ErrorMessage });
        }


        [HttpPost]
        public JsonResult Report(string ID, string No, bool IsToPDF)
        {
            BaseResult result;
            string FileName;
            if (IsToPDF)
            {
                //result = _WaterFastnessService.ToPdfWaterFastnessDetail(ID, No, out FileName, false);
                result = _WaterFastnessService.ToReport(ID, out FileName, true, false);
            }
            else
            {
                //result = _WaterFastnessService.ToExcelWaterFastnessDetail(ID, No, out FileName, false);
                result = _WaterFastnessService.ToReport(ID, out FileName, false, false);
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
    }
}