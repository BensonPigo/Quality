using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class FabricColorFastnessController : BaseController
    {
        private IFabricColorFastness_Service _FabricColorFastness_Service;

        public FabricColorFastnessController()
        {
            _FabricColorFastness_Service = new FabricColorFastness_Service();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.FabricColorFastness,,";
        }

        public ActionResult Index()
        {
            FabricColorFastness_ViewModel model = new FabricColorFastness_ViewModel()
            {
                ColorFastness_MainList = new List<ColorFastness_Result>()
            };

            return View(model);
        }

        public ActionResult IndexBack(string PoID)
        {
            FabricColorFastness_ViewModel model = _FabricColorFastness_Service.Get_Main(PoID);
            ViewBag.QueryPoID = PoID;
            return View("Index", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(string QueryPoID)
        {
            // 21051739BB
            FabricColorFastness_ViewModel model = _FabricColorFastness_Service.Get_Main(QueryPoID);

            if (!model.Result)
            {
                model.ColorFastness_MainList = new List<ColorFastness_Result>();
                model.ErrorMessage = $"msg.WithInfo('" + model.ErrorMessage.ToString().Replace("\r\n", "<br />") + "');";
            }

            ViewBag.QueryPoID = QueryPoID;
            return View("Index", model);
        }

        [HttpPost]
        public JsonResult SaveMaster(FabricColorFastness_ViewModel Main)
        {
            var result = _FabricColorFastness_Service.Save_ColorFastness_1stPage(Main.PoID, Main.ColorFastnessLaboratoryRemark);
            ViewBag.QueryPoID = Main.PoID;
            return Json(result);
        }

        public JsonResult MainDetailDelete(string ID)
        {
            BaseResult result = _FabricColorFastness_Service.DeleteColorFastness(ID);

            return Json(result);
        }

        public ActionResult Detail(string PoID, string ID, string EditMode)
        {
            FabricColorFastness_ViewModel FabricColorFastnessModel = new FabricColorFastness_ViewModel();
            Fabric_ColorFastness_Detail_ViewModel model = new Fabric_ColorFastness_Detail_ViewModel();
            if (Convert.ToBoolean(EditMode) && string.IsNullOrEmpty(ID))
            {
                model.Main = new ColorFastness_Result();
                model.Detail = new List<Fabric_ColorFastness_Detail_Result>();
                model.Main.Status = "New";
                model.Main.POID = PoID;
            }
            else
            {
                model = _FabricColorFastness_Service.GetDetailBody(ID);
            }

            if (TempData["Model"] != null)
            {
                Fabric_ColorFastness_Detail_ViewModel saveResult = (Fabric_ColorFastness_Detail_ViewModel)TempData["Model"];
                model.Main.InspDate = saveResult.Main.InspDate;
                model.Main.Article = saveResult.Main.Article;
                model.Main.Inspector = saveResult.Main.Inspector;
                model.Main.InspectionName = saveResult.Main.InspectionName;
                model.Main.Cycle = saveResult.Main.Cycle;
                model.Main.Detergent = saveResult.Main.Detergent;
                model.Main.Machine = saveResult.Main.Machine;
                model.Main.Drying = saveResult.Main.Drying;
                model.Main.Remark = saveResult.Main.Remark;
                model.Detail = saveResult.Detail;
                model.Result = saveResult.Result;
                model.ErrorMessage = $"msg.WithInfo('{saveResult.ErrorMessage.ToString().Replace("\r\n", "<br />") }');";
            }

            List<string> Scales = _FabricColorFastness_Service.Get_Scales();
            List<SelectListItem> ScalesList = new SetListItem().ItemListBinding(Scales);
            ViewBag.Temperature_List = FabricColorFastnessModel.Temperature_List;
            ViewBag.Cycle_List = FabricColorFastnessModel.Cycle_List;
            ViewBag.Detergent_List = FabricColorFastnessModel.Detergent_List;
            ViewBag.Machine_List = FabricColorFastnessModel.Machine_List;
            ViewBag.Drying_List = FabricColorFastnessModel.Drying_List;
            ViewBag.ChangeScaleList = ScalesList;
            ViewBag.ResultChangeList = FabricColorFastnessModel.Result_Source;
            ViewBag.StainingScaleList = ScalesList;
            ViewBag.ResultStainList = FabricColorFastnessModel.Result_Source;
            ViewBag.EditMode = EditMode;
            ViewBag.FactoryID = this.FactoryID;
            return View(model);
        }

        [HttpPost]
        public ActionResult AddDetailRow(int lastNO, string GroupNO)
        {
            FabricColorFastness_ViewModel FabricColorFastnessModel = new FabricColorFastness_ViewModel();
            List<string> Scales = _FabricColorFastness_Service.Get_Scales();


            int i = lastNO;
            string html = "";
            html += "<tr>";
            html += "<td><input id='Seq" + i + "' idx=" + i + " type ='hidden'></input><input id='Detail_" + i + "__SubmitDate' name='Detail[" + i + "].SubmitDate' class='form-control date-picker' type='text' value=''></td>"; //SubmitDate
            html += "<td><input id='Detail_" + i + "__ColorFastnessGroup' name='Detail[" + i + "].ColorFastnessGroup' type='number' max='99' maxlength='2' min='0' step='1' oninput='value=OvenGroupCheck(value)' value='" + GroupNO + "'></td>"; // ColorFastnessGroup
            html += "<td style='width: 11vw;'><div style='width:10vw;'><input id='Detail_" + i + "__Seq' name='Detail[" + i + "].Seq' idv='" + i.ToString() + "' class ='InputDetailSEQSelectItem' type='text'  style = 'width: 6vw'> <input id='btnDetailSEQSelectItem'  idv='" + i.ToString() + "' type='button' class='btnDetailSEQSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // seq
            html += "<td style='width: 11vw;'><div style='width:10vw;'><input id='Detail_" + i + "__Roll' name='Detail[" + i + "].Roll' idv='" + i.ToString() + "' class ='InputDetailRollSelectItem' type='text' style = 'width: 6vw'> <input id='btnDetailRollSelectItem' idv='" + i.ToString() + "' type='button' class='btnDetailRollSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // roll
            html += "<td><input id='Detail_" + i + "__Dyelot' name='Detail[" + i + "].Dyelot' type='text' readonly='readonly'></td>"; // Dyelot
            html += "<td><input id='Detail_" + i + "__Refno' name='Detail[" + i + "].Refno' type='text' readonly='readonly'></td>"; // Refno
            html += "<td><input id='Detail_" + i + "__SCIRefno' name='Detail[" + i + "].SCIRefno' type='text' readonly='readonly'></td>"; // SCIRefno
            html += "<td><input id='Detail_" + i + "__ColorID' name='Detail[" + i + "].SCIRefno' type='text' readonly='readonly'></td>"; // ColorID
            html += "<td><input id='Detail_" + i + "__Result' name='Detail[" + i + "].Result' type='text' readonly='readonly' class='blue'></td>"; // Result

            html += "<td><select id='Detail_" + i + "__changeScale' name='Detail[" + i + "].changeScale'><option value=''></option>"; // changeScale
            foreach (string val in Scales)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Detail_" + i + "__ResultChange' name='Detail[" + i + "].ResultChange' onchange='selectChange(this)'><option value=''></option>"; // ResultChange
            foreach (var val in FabricColorFastnessModel.Result_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Detail_" + i + "__StainingScale' name='Detail[" + i + "].StainingScale'><option value=''></option>"; // StainingScale
            foreach (string val in Scales)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select id='Detail_" + i + "__ResultStain' name='Detail[" + i + "].ResultStain' onchange='selectChange(this)'><option value=''></option>"; // ResultStain
            foreach (var val in FabricColorFastnessModel.Result_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            html += "<td><input id='Detail_" + i + "__Remark' name='Detail[" + i + "].Remark' type='text'></td>"; // remark

            html += "<td></td>"; // LastUpdate

            html += "<td><img class='detailDelete display-None' src='/Image/Icon/Delete.png' width='30'></td>";
            html += "</tr>";

            return Content(html);
        }

        [HttpPost]
        public ActionResult DetailSave(Fabric_ColorFastness_Detail_ViewModel req)
        {
            if (req.Detail == null)
            {
                req.Detail = new List<Fabric_ColorFastness_Detail_Result>();
            }

            BaseResult result = _FabricColorFastness_Service.Save_ColorFastness_2ndPage(req, this.MDivisionID, this.UserID);
            if (result.Result)
            {
                // 找地方寫入 new ID
                req.Main.ID = string.IsNullOrEmpty(result.ErrorMessage) ? req.Main.ID : result.ErrorMessage;
                return RedirectToAction("Detail", new { POID = req.Main.POID, ID = req.Main.ID, EditMode = false });
            }

            req.Result = result.Result;
            req.ErrorMessage = result.ErrorMessage;
            TempData["Model"] = req;
            return RedirectToAction("Detail", new { POID = req.Main.POID, ID = req.Main.ID, EditMode = false });
        }

        [HttpPost]
        public JsonResult Encode_Detail(string ID, FabricColorFastness_Service.DetailStatus Type)
        {
            string ColorFastnessResult = string.Empty;
            // BaseResult result = _FabricColorFastness_Service.Encode_ColorFastness(ID, Type, this.UserID, out ColorFastnessResult);
            BaseResult result = new BaseResult();
            result.Result = true;
            ColorFastnessResult = "Fail";
            return Json(new { result.Result, result.ErrorMessage, ColorFastnessResult });
        }

        [HttpPost]
        public JsonResult FailMail(string POID, string ID, string TO, string CC)
        {
            BaseResult result = _FabricColorFastness_Service.SentMail(POID, ID, TO, CC);
            return Json(result);
        }

        [HttpPost]
        public JsonResult Report(string ID, bool IsToPDF)
        {
            Fabric_ColorFastness_Detail_ViewModel result;
            if (IsToPDF)
            {
                result = _FabricColorFastness_Service.ToPDF(ID, false);
            }
            else
            {
                result = _FabricColorFastness_Service.ToExcel(ID, false);
            }

            string FileName = result.reportPath;
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
    }
}