using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject;
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


        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(string QueryPoID)
        {
            // 21051739BB
            FabricColorFastness_ViewModel model = _FabricColorFastness_Service.Get_Main(QueryPoID);
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

        public ActionResult IndexBack(string PoID)
        {
            FabricColorFastness_ViewModel model = _FabricColorFastness_Service.Get_Main(PoID);
            ViewBag.QueryPoID = PoID;
            return View("Index", model);
        }

        public JsonResult MainDetailDelete(string ID, string No)
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
        public ActionResult AddDetailRow(string POID, int lastNO)
        {
            FabricColorFastness_ViewModel FabricColorFastnessModel = new FabricColorFastness_ViewModel();
            List<string> Scales = _FabricColorFastness_Service.Get_Scales();


            int i = lastNO;
            string html = "";
            html += "<tr>";
            html += "<td> <input id ='Seq' idx= " + i + " type ='hidden'></input> <input class='form-control date-picker' type='text' value=''></td>"; //SubmitDate
            html += "<td><input type='text'></td>"; // ColorFastnessGroup
            html += "<td style='width: 11vw;'><div style='width:10vw;'><input id='Detail_" + i + "__Seq' idv='" + i.ToString() + "' class ='InputDetailSEQSelectItem' type='text'  style = 'width: 6vw'> <input id='btnDetailSEQSelectItem'  idv='" + i.ToString() + "' type='button' class='btnDetailSEQSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // seq
            html += "<td style='width: 11vw;'><div style='width:10vw;'><input id='Detail_" + i + "__Roll'idv='" + i.ToString() + "' class ='InputDetailRollSelectItem' type='text' style = 'width: 6vw'> <input id='btnDetailRollSelectItem' idv='" + i.ToString() + "' type='button' class='btnDetailRollSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // roll
            html += "<td><input id='Detail_" + i + "__Dyelot' type='text' readonly='readonly'></td>"; // Dyelot
            html += "<td><input id='Detail_" + i + "__Refno' type='text' readonly='readonly'></td>"; // Refno
            html += "<td><input id='Detail_" + i + "__SCIRefno' type='text' readonly='readonly'></td>"; // SCIRefno
            html += "<td><input id='Detail_" + i + "__ColorID' type='text' readonly='readonly'></td>"; // ColorID
            html += "<td><input  readonly='readonly'  id='Detail_" + i + "__Result' name='Detail[" + i + "].Result'  class='detailResultColor' type='text'></td>"; // Result

            html += "<td><select ><option value=''></option>"; // changeScale
            foreach (string val in Scales)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Detail_" + i + "__ResultChange' name='Detail[" + i + "].ResultChange' ><option value=''></option>"; // ResultChange
            foreach (var val in FabricColorFastnessModel.Result_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            html += "<td><select ><option value=''></option>"; // StainingScale
            foreach (string val in Scales)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Detail_" + i + "__ResultStain' name='Detail[" + i + "].ResultStain' ><option value=''></option>"; // ResultStain
            foreach (var val in FabricColorFastnessModel.Result_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            html += "<td><input type='text'></td>"; // remark

            html += "<td></td>"; // LastUpdate

            html += "<td><img  class='detailDelete' src='/Image/Icon/Delete.png' width='30'></td>";
            html += "</tr>";

            return Content(html);
        }


        [HttpPost]
        public JsonResult Encode_Detail(string POID, string TestNo)
        {
            string ColorFastnessResult = string.Empty;
            BaseResult result = new BaseResult();
            result.Result = true;
            ColorFastnessResult = "Fail";
            return Json(new { result.Result, result.ErrorMessage, ColorFastnessResult });
        }

    }
}