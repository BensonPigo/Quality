using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.BulkFGT;
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
    public class PerspirationFastnessController : BaseController
    {
        public PerspirationFastnessService _PerspirationFastnessService;
        private List<string> Results = new List<string>() { "Pass", "Fail" };
        private List<string> Temperatures = new List<string>() { "35", "36", "37", "38", "39" };
        private List<string> Times = new List<string>() { "4", "8", "12", "24" };
        private List<string> MetalContents = new List<string>() { "None", "Metail Printing", "Metal Thread" };

        public PerspirationFastnessController()
        {
            _PerspirationFastnessService = new PerspirationFastnessService();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.PerspirationFastness,,";
        }

        // GET: BulkFGT/PerspirationFastness
        public ActionResult Index()
        {
            PerspirationFastness_Result model = new PerspirationFastness_Result()
            {
                Main = new PerspirationFastness_Main(),
                Details = new List<PerspirationFastness_Detail>()
            };

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(string POID)
        {
            // 21051739BB
            PerspirationFastness_Result model = _PerspirationFastnessService.GetPerspirationFastness_Result(POID);
            if (!model.Result)
            {
                model = new PerspirationFastness_Result()
                {
                    Main = new PerspirationFastness_Main(),
                    Details = new List<PerspirationFastness_Detail>(),
                    ErrorMessage = $@"msg.WithInfo(""{ (string.IsNullOrEmpty(model.ErrorMessage) ? string.Empty : model.ErrorMessage.Replace("\r\n", "<br />"))  }"");",
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
            PerspirationFastness_Result model = _PerspirationFastnessService.GetPerspirationFastness_Result(POID);
            ViewBag.POID = POID;
            return View("Index", model);
        }
        public ActionResult Detail(string POID, string TestNo, string EditMode, string BrandID)
        {
            PerspirationFastness_Detail_Result model = _PerspirationFastnessService.GetPerspirationFastness_Detail_Result(POID, TestNo, BrandID);

            List<SelectListItem> ScaleIDList = new SetListItem().ItemListBinding(model.ScaleIDs);
            List<SelectListItem> ResultChangeList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> ResultStainList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> TimeList = new SetListItem().ItemListBinding(Times);
            List<SelectListItem> MetalContentList = new SetListItem().ItemListBinding(MetalContents);

            if (Convert.ToBoolean(EditMode) && string.IsNullOrEmpty(TestNo))
            {
                model.Main.Status = "New";
            }

            if (TempData["ModelPerspirationFastness"] != null)
            {
                PerspirationFastness_Detail_Result saveResult = (PerspirationFastness_Detail_Result)TempData["ModelPerspirationFastness"];
                model.Main.InspDate = saveResult.Main.InspDate;
                model.Main.Article = saveResult.Main.Article;
                model.Main.Inspector = saveResult.Main.Inspector;
                model.Main.InspectorName = saveResult.Main.InspectorName;
                model.Main.Remark = saveResult.Main.Remark;
                model.Details = saveResult.Details;
                model.Result = saveResult.Result;
                model.ErrorMessage = $@"msg.WithInfo('{(string.IsNullOrEmpty(saveResult.ErrorMessage) ? string.Empty : saveResult.ErrorMessage.Replace("'",string.Empty)) }');EditMode=true;";
                EditMode = "True";
            }

            ViewBag.ChangeScaleList = ScaleIDList;
            ViewBag.ResultChangeList = ResultChangeList;
            ViewBag.StainingScaleList = ScaleIDList;
            ViewBag.ResultStainList = ResultStainList;
            ViewBag.TemperatureList = TemperatureList;
            ViewBag.TimeList = TimeList;
            ViewBag.EditMode = EditMode;
            ViewBag.MetalContentList = MetalContentList;
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.ErrorMessage = string.Empty;
            ViewBag.UserMail = this.UserMail;
            return View(model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult DetailSave(PerspirationFastness_Detail_Result req)
        {
            req.MDivisionID = this.MDivisionID;

            req.Main.TestBeforePicture = req.Main.TestBeforePicture == null ? null : ImageHelper.ImageCompress(req.Main.TestBeforePicture);
            req.Main.TestAfterPicture = req.Main.TestAfterPicture == null ? null : ImageHelper.ImageCompress(req.Main.TestAfterPicture);

            BaseResult result = _PerspirationFastnessService.SavePerspirationFastnessDetail(req, this.UserID);
            if (result.Result)
            {
                // 找地方寫入 TestNo
                req.Main.TestNo = string.IsNullOrEmpty(result.ErrorMessage) ? req.Main.TestNo : result.ErrorMessage;
                return RedirectToAction("Detail", new { POID = req.Main.POID, TestNo = req.Main.TestNo, EditMode = false });
            }

            req.Result = result.Result;
            req.ErrorMessage = result.ErrorMessage;
            TempData["ModelPerspirationFastness"] = req;
            ViewBag.UserMail = this.UserMail;
            return RedirectToAction("Detail", new { POID = req.Main.POID, TestNo = req.Main.TestNo, EditMode = false });
        }
        public JsonResult MainDetailDelete(string ID, string No)
        {
            BaseResult result = _PerspirationFastnessService.DeletePerspirationFastness(ID, No);

            return Json(result);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult SaveMaster(PerspirationFastness_Main Main)
        {
            var result = _PerspirationFastnessService.SavePerspirationFastnessMain(Main);

            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult AddDetailRow(string POID, int lastNO, string GroupNO ,string BrandID)
        {
            PerspirationFastness_Detail_Result model = _PerspirationFastnessService.GetPerspirationFastness_Detail_Result(POID, "", BrandID);


            List<SelectListItem> ResultChangeList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> ResultStainList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> TimeList = new SetListItem().ItemListBinding(Times);

            int i = lastNO;
            PerspirationFastness_Detail_Detail detail = new PerspirationFastness_Detail_Detail();
            string html = "";
            html += "<tr idx='" + i + "'>";
            html += "<td><input id='Seq" + i + "' idx='" + i + "' type ='hidden'></input> <input id='Details_" + i + "__SubmitDate' name='Details[" + i + "].SubmitDate' class='form-control date-picker width9vw' type='text' value=''></td>";
            html += "<td><input id='Details_" + i + "__PerspirationFastnessGroup' name='Details[" + i + "].PerspirationFastnessGroup' type='number' max='99' maxlength='2' min='0' step='1' oninput='value=PerspirationFastnessGroupCheck(value)' value='" + GroupNO + "'></td>"; // group
            html += "<td><div style='width:10vw;'><input id='Details_" + i + "__SEQ' name='Details[" + i + "].SEQ' idv='" + i.ToString() + "' class ='InputDetailSEQSelectItem' type='text'  style = 'width: 6vw'> <input id='btnDetailSEQSelectItem'  idv='" + i.ToString() + "' type='button' class='btnDetailSEQSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // seq
            html += "<td><div style='width:10vw;'><input id='Details_" + i + "__Roll' name='Details[" + i + "].Roll' idv='" + i.ToString() + "' class ='InputDetailRollSelectItem' type='text' style = 'width: 6vw'> <input id='btnDetailRollSelectItem' idv='" + i.ToString() + "' type='button' class='btnDetailRollSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // roll
            html += "<td><input id='Details_" + i + "__Dyelot' name='Details[" + i + "].Dyelot' type='text' readonly='readonly'></td>"; // dyelot
            html += "<td><input id='Details_" + i + "__Refno' name='Details[" + i + "].Refno' type='text' readonly='readonly'></td>"; // Refno
            html += "<td><input id='Details_" + i + "__SCIRefno' name='Details[" + i + "].SCIRefno' type='text' readonly='readonly'></td>"; // SCIRefno
            html += "<td><input id='Details_" + i + "__ColorID' name='Details[" + i + "].ColorID' type='text' readonly='readonly'></td>"; // ColorID
            html += "<td><input  readonly='readonly'  id='Details_" + i + "__Result' name='Details[" + i + "].Result'  class='blue width6vw' type='text'></td>"; // Result

            html += @"<td class=""Alkaline""><select id='Details_" + i + "__AlkalineChangeScale' name='Details[" + i + "].AlkalineChangeScale'>"; // ChangeScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AlkalineResultChange' name='Details[" + i + "].AlkalineResultChange' >"; // ResultAcetate
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select id='Details_" + i + "__AlkalineAcetateScale' name='Details[" + i + "].AlkalineAcetateScale'>"; // AcetateScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AlkalineResultAcetate' name='Details[" + i + "].AlkalineResultAcetate' >"; // ResultAcetate
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select id='Details_" + i + "__AlkalineCottonScale' name='Details[" + i + "].AlkalineCottonScale'>"; // CottonScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AlkalineResultCotton' name='Details[" + i + "].AlkalineResultCotton' >"; // ResultCotton
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select id='Details_" + i + "__AlkalineNylonScale' name='Details[" + i + "].AlkalineNylonScale'>"; // NylonScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AlkalineResultNylon' name='Details[" + i + "].AlkalineResultNylon' >"; // ResultNylon
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select id='Details_" + i + "__AlkalinePolyesterScale' name='Details[" + i + "].AlkalinePolyesterScale'>"; // PolyesterScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AlkalineResultPolyester' name='Details[" + i + "].AlkalineResultPolyester' >"; // ResultPolyester
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select id='Details_" + i + "__AlkalineAcrylicScale' name='Details[" + i + "].AlkalineAcrylicScale'>"; // AcrylicScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AlkalineResultAcrylic' name='Details[" + i + "].AlkalineResultAcrylic' >"; // ResultAcrylic
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select id='Details_" + i + "__AlkalineWoolScale' name='Details[" + i + "].AlkalineWoolScale'>"; // WoolScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Alkaline""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AlkalineResultWool' name='Details[" + i + "].AlkalineResultWool' >"; // ResultWool
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";


            html += @"<td class=""Acid""><select id='Details_" + i + "__AcidChangeScale' name='Details[" + i + "].AcidChangeScale'>"; // ChangeScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AcidResultChange' name='Details[" + i + "].AcidResultChange' >"; // ResultAcetate
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select id='Details_" + i + "__AcidAcetateScale' name='Details[" + i + "].AcidAcetateScale'>"; // AcetateScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AcidResultAcetate' name='Details[" + i + "].AcidResultAcetate' >"; // ResultAcetate
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select id='Details_" + i + "__AcidCottonScale' name='Details[" + i + "].AcidCottonScale'>"; // CottonScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AcidResultCotton' name='Details[" + i + "].AcidResultCotton' >"; // ResultCotton
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select id='Details_" + i + "__AcidNylonScale' name='Details[" + i + "].AcidNylonScale'>"; // NylonScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AcidResultNylon' name='Details[" + i + "].AcidResultNylon' >"; // ResultNylon
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select id='Details_" + i + "__AcidPolyesterScale' name='Details[" + i + "].AcidPolyesterScale'>"; // PolyesterScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AcidResultPolyester' name='Details[" + i + "].AcidResultPolyester' >"; // ResultPolyester
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select id='Details_" + i + "__AcidAcrylicScale' name='Details[" + i + "].AcidAcrylicScale'>"; // AcrylicScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AcidResultAcrylic' name='Details[" + i + "].AcidResultAcrylic' >"; // ResultAcrylic
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select id='Details_" + i + "__AcidWoolScale' name='Details[" + i + "].AcidWoolScale'>"; // WoolScale
            foreach (string val in model.ScaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += @"<td class=""Acid""><select onchange='selectChange(this)' class='blue' id='Details_" + i + "__AcidResultWool' name='Details[" + i + "].AcidResultWool' >"; // ResultWool
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
        [SessionAuthorizeAttribute]
        public JsonResult Encode_Detail(string POID, string TestNo)
        {
            string PerspirationFastnessResult = string.Empty;
            BaseResult result = _PerspirationFastnessService.EncodePerspirationFastnessDetail(POID, TestNo, out PerspirationFastnessResult);
            return Json(new { result.Result, result.ErrorMessage, PerspirationFastnessResult });
        }

        //[HttpPost]
        //public JsonResult FailMail(string ID, string No, string TO, string CC)
        //{
        //    SendMail_Result result = _PerspirationFastnessService.SendMail(TO, CC, ID, No, false);
        //    return Json(result);
        //}
        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Amend_Detail(string POID, string TestNo)
        {
            BaseResult result = _PerspirationFastnessService.AmendPerspirationFastnessDetail(POID, TestNo);
            return Json(new { result.Result, ErrorMessage = (string.IsNullOrEmpty(result.ErrorMessage) ? string.Empty : result.ErrorMessage.Replace("'", string.Empty)) });
        }


        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Report(string ID, string No, bool IsToPDF)
        {
            BaseResult result;
            string FileName;
            if (IsToPDF)
            {
                //result = _PerspirationFastnessService.ToPdfPerspirationFastnessDetail(ID, No, out FileName, false);
                result = _PerspirationFastnessService.ToReport(ID, out FileName, true, false);
            }
            else
            {
                //result = _PerspirationFastnessService.ToExcelPerspirationFastnessDetail(ID, No, out FileName, false);
                result = _PerspirationFastnessService.ToReport(ID, out FileName, false, false);
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult SendMail(string ID, string No, string TO, string CC)
        {
            SendMail_Result result = _PerspirationFastnessService.SendMail(TO, CC, ID, No, false);
            return Json(result);
            //this.CheckSession();

            //BaseResult result = null;
            //string FileName = string.Empty;

            //result = _PerspirationFastnessService.ToReport(ID, out FileName, true, false);
            //if (!result.Result)
            //{
            //    result.ErrorMessage = result.ErrorMessage.ToString();
            //}
            //string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;

            //return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, reportPath = reportPath, FileName = FileName });
        }
    }
}