using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
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
        public ActionResult Index(string POID)
        {

            FabricOvenTest_Result model = _FabricOvenTestService.GetFabricOvenTest_Result("21051739BB");

            ViewBag.POID = POID;
            return View(model);
        }

        public ActionResult IndexBack(string POID)
        {

            FabricOvenTest_Result model = _FabricOvenTestService.GetFabricOvenTest_Result("21051739BB");

            ViewBag.POID = POID;
            return View("Index",model);
        }

        public ActionResult Detail(string POID, string TestNo,string EditMode)
        {

            FabricOvenTest_Detail_Result model = _FabricOvenTestService.GetFabricOvenTest_Detail_Result("21051739BB", TestNo);

            List<SelectListItem> ScaleIDList = new SetListItem().ItemListBinding(model.ScaleIDs);
            List<SelectListItem> ResultChangeList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> ResultStainList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> TimeList = new SetListItem().ItemListBinding(Times);

            if (Convert.ToBoolean(EditMode) && string.IsNullOrEmpty(TestNo))
            {
                model.Main.Status = "New";
            }
            ViewBag.ChangeScaleList = ScaleIDList;
            ViewBag.ResultChangeList = ResultChangeList;
            ViewBag.StainingScaleList = ScaleIDList;
            ViewBag.ResultStainList = ResultStainList;
            ViewBag.TemperatureList = TemperatureList;
            ViewBag.TimeList = TimeList;
            ViewBag.EditMode = EditMode;
            return View(model);
        }



        [HttpPost]
        public JsonResult SaveMaster(FabricOvenTest_Main Main)
        {
            var result = _FabricOvenTestService.SaveFabricOvenTestMain(Main);
           


            return Json(result);


        }

        [HttpPost]
        public ActionResult AddDetailRow(string POID, int lastNO)
        {
            FabricOvenTest_Detail_Result model = _FabricOvenTestService.GetFabricOvenTest_Detail_Result("21051739BB", "");


            List<SelectListItem> ResultChangeList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> ResultStainList = new SetListItem().ItemListBinding(Results);
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> TimeList = new SetListItem().ItemListBinding(Times);

            int i = lastNO;
            FabricOvenTest_Detail_Detail detail = new FabricOvenTest_Detail_Detail();
            string html = "";
            html += "<tr>";
            html += "<td> <input id ='Seq' idx= " + i + " type ='hidden'></input> <input class='form-control date-picker' type='text' value=''></td>";
            html += "<td><input type='text'></td>"; // group
            html += "<td style='width: 10vw;'><div style='width:9vw;'><input id='Details_" + i + "__SEQ' type='text' readonly='readonly' style = 'width: 6vw'><input id='btnDetailSEQSelectItem'  idv='" + i.ToString() + "' type='button' class='btnDetailSEQSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // seq
            html += "<td style='width: 10vw;'><div style='width:9vw;'><input id='Details_" + i + "__Roll' type='text' readonly='readonly' style = 'width: 6vw'><input id='btnDetailRollSelectItem' idv='" + i.ToString() + "' type='button' class='btnDetailRollSelectItem OnlyEdit site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>"; // roll
            html += "<td><input id='Details_" + i + "__Dyelot' type='text' readonly='readonly'></td>"; // dyelot
            html += "<td><input id='Details_" + i + "__Refno' type='text' readonly='readonly'></td>"; // Refno
            html += "<td><input id='Details_" + i + "__SCIRefno' type='text' readonly='readonly'></td>"; // SCIRefno
            html += "<td><input id='Details_" + i + "__ColorID' type='text' readonly='readonly'></td>"; // ColorID
            html += "<td><input  readonly='readonly'  id='Details_" + i + "__Result' name='Details[" + i + "].Result'  class='detailResultColor' type='text'></td>"; // Result

            html += "<td><select ><option value=''></option>"; // ChangeScale
            foreach (string val in model.ScaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultChange' name='Details[" + i + "].ResultChange' ><option value=''></option>"; // ResultChange
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select ><option value=''></option>"; // StainingScale
            foreach (string val in model.ScaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><select onchange='selectChange(this)' id='Details_" + i + "__ResultStain' name='Details[" + i + "].ResultStain' ><option value=''></option>"; // ResultStain
            foreach (string val in Results)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><input type='text'></td>"; // remark

            html += "<td><input type='text'></td>"; // LastUpdate


            html += "<td><select ><option value=''></option>"; // Temperature
            foreach (string val in Temperatures)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";


            html += "<td><select ><option value=''></option>"; // Time
            foreach (string val in Times)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";

            html += "<td><img class='detailDelete display-None' src='/Image/Icon/Delete.png' width='30'></td>";
            html += "</tr>";

            return Content(html);
        }

    }
}