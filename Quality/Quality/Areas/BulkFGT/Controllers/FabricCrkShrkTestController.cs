using BusinessLogicLayer.Interface;
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
    public class FabricCrkShrkTestController : BaseController
    {
        public FabricCrkShrkTestController()
        {
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.FabricCrkShrkTest,,";
        }

        // GET: BulkFGT/FabricCrkShrkTest
        public ActionResult Index()
        {
            DatabaseObject.ResultModel.FabricCrkShrkTest_Result fabricCrkShrkTest_Result = new DatabaseObject.ResultModel.FabricCrkShrkTest_Result()
            {
                Main = new DatabaseObject.ResultModel.FabricCrkShrkTest_Main(),
                Details = new List<DatabaseObject.ResultModel.FabricCrkShrkTest_Detail>()
            };

            return View(fabricCrkShrkTest_Result);
        }

        [HttpPost]
        public ActionResult Index(string POID)
        {
            ViewBag.POID = POID;
            IFabricCrkShrkTest_Service fabricOvenTestService = new FabricCrkShrkTest_Service();
            DatabaseObject.ResultModel.FabricCrkShrkTest_Result fabricCrkShrkTest_Result =  fabricOvenTestService.GetFabricCrkShrkTest_Result(POID);
            return View(fabricCrkShrkTest_Result);
        }

        [HttpPost]
        public ActionResult SaveIndex(DatabaseObject.ResultModel.FabricCrkShrkTest_Main main, List<DatabaseObject.ResultModel.FabricCrkShrkTest_Detail> detail)
        {
            return Json(main);
        }

        [HttpGet]
        public ActionResult IndexBack(string POID)
        {
            ViewBag.POID = POID;
            IFabricCrkShrkTest_Service fabricOvenTestService = new FabricCrkShrkTest_Service();
            DatabaseObject.ResultModel.FabricCrkShrkTest_Result fabricCrkShrkTest_Result = fabricOvenTestService.GetFabricCrkShrkTest_Result(POID);
            return View("Index", fabricCrkShrkTest_Result);
        }

        [HttpPost]
        public ActionResult SaveMain()
        {
            return View();
        }

        public ActionResult CrockingTest(long ID)
        {
            DatabaseObject.ResultModel.FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result = new DatabaseObject.ResultModel.FabricCrkShrkTestCrocking_Result()
            {
                ID = ID,
                ScaleIDs = new List<string>() { "a", "b" },
                CrockingTestOption = 1, // 0 會隱藏明細欄位
                Crocking_Main = new DatabaseObject.ResultModel.FabricCrkShrkTestCrocking_Main
                {
                    Crocking = "Fail",
                    NonCrocking = true,
                    POID = "21051318BB",
                    SEQ = "01 01",
                    ColorID = "1",
                    ArriveQty = 2,
                    WhseArrival = System.DateTime.Now,
                    ExportID = "3",
                    Supp = "4",
                    CrockingDate = System.DateTime.Now.AddDays(-1),
                    StyleID = "5",
                    SCIRefno = "6",
                    Name = "7",
                    BrandID = "8",
                    Refno = "9",
                    DescDetail = "10",
                    CrockingRemark = "11",
                    CrockingEncdoe = false
                },
                Crocking_Detail = new List<DatabaseObject.ResultModel.FabricCrkShrkTestCrocking_Detail>(),
            };

            FabricCrkShrkTestCrocking_Detail test1 = new FabricCrkShrkTestCrocking_Detail()
            {
                Roll = "1",
                Dyelot = "2",
                Result = "Fail",
                DryScale = "a",
                ResultDry = "Fail",
                DryScale_Weft = "b",
                ResultDry_Weft = "Pass",
                WetScale = "a",
                ResultWet = "Fail",
                WetScale_Weft = "b",
                ResultWet_Weft = "Fail",
                Inspdate = System.DateTime.Now,
                Inspector = "3",
                Name = "3",
                Remark = "3",
                LastUpdate = "3"
            };
            fabricCrkShrkTestCrocking_Result.Crocking_Detail.Add(test1);

            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };

            ViewBag.ResultList= new SetListItem().ItemListBinding(resultType);
            ViewBag.ScaleIDsList = new SetListItem().ItemListBinding(fabricCrkShrkTestCrocking_Result.ScaleIDs);

            return View(fabricCrkShrkTestCrocking_Result);
        }

        [HttpPost]
        public ActionResult AddDetailRow(int lastNO)
        {
            lastNO = lastNO + 1;

            // 假資料，因為清單來源目前沒有            
            string testTemp = "<option value='a'>a</option ><option value='b'>b</option>";

            string html = "";
            html += $"<tr idx='{lastNO}' class='row-content' style='vertical-align:middle;text-align:center;'>";
            html += "<td>";
            html += $"<input id='Crocking_Detail_{lastNO}__Roll' name='Crocking_Detail[{lastNO}].Roll' class='inputRollSelectItem' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Crocking_Detail_{lastNO}__Dyelot' name='Crocking_Detail[{lastNO}].Dyelot' class='inputRollSelectItem' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Crocking_Detail_{lastNO}__Result' name='Crocking_Detail[{lastNO}].Result' readonly='readonly' style='color:blue' type='text' value='Pass'>";
            html += "</td>";
            html += "<td style='color:blue'>";
            html += $"<select id='Crocking_Detail_{lastNO}__DryScale' name='Crocking_Detail[{lastNO}].DryScale' style='width:157px'>";

            html += testTemp;

            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select class='resultList result{lastNO}' id='Crocking_Detail_{lastNO}__ResultDry' name='Crocking_Detail[{lastNO}].ResultDry' onchange='changeResultColor(this)' style='width:157px'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Crocking_Detail_{lastNO}__DryScale_Weft' name='Crocking_Detail[{lastNO}].DryScale_Weft' style='width:157px'>";

            html += testTemp;

            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select class='resultList result{lastNO}' id='Crocking_Detail_{lastNO}__ResultDry_Weft' name='Crocking_Detail[{lastNO}].ResultDry_Weft' onchange='changeResultColor(this)' style='width:157px'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Crocking_Detail_{lastNO}__WetScale' name='Crocking_Detail[{lastNO}].WetScale' style='width:157px'>";

            html += testTemp;

            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select class='resultList result{lastNO}' id='Crocking_Detail_{lastNO}__ResultWet' name='Crocking_Detail[{lastNO}].ResultWet' onchange='changeResultColor(this)' style='width:157px'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Crocking_Detail_{lastNO}__WetScale_Weft' name='Crocking_Detail[{lastNO}].WetScale_Weft' style='width:157px'>";

            html += testTemp;

            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select class='resultList result{lastNO}' id='Crocking_Detail_{lastNO}__ResultWet_Weft' name='Crocking_Detail[{lastNO}].ResultWet_Weft' onchange='changeResultColor(this)' style='width:157px'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='date-picker hasDatepicker' data-val='true' data-val-date='欄位 Inspdate 必須是日期。' id='Crocking_Detail_{lastNO}__Inspdate' name='Crocking_Detail[{lastNO}].Inspdate' type='text' value='2021/09/06' disabled='disabled'>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Crocking_Detail_{lastNO}__Inspector' name='Crocking_Detail[{lastNO}].Inspector' type='text' value=''>";
            html += $"<input id='btnDetailInspectorSelectItem' type='button' class='site-btn btn-blue' style='margin:0;border:0;' value='...' disabled='disabled'>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Crocking_Detail_{lastNO}__Name' name='Crocking_Detail[{lastNO}].Name' readonly=readonly type='text' value=''>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Crocking_Detail_{lastNO}__Remark' name='Crocking_Detail[{lastNO}].Remark' readonly=readonly type='text' value=''>";
            html += "</td>";
            html += "<td>";// last date
            html += "</td>";
            html += "<td><img class='detailDelete' src='/Image/Icon/Delete.png' width='30'></td>";
            html += "</tr>";

            return Content(html);
        }

        [HttpPost]
        public ActionResult CrockingTest(DatabaseObject.ResultModel.FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result)
        {
            fabricCrkShrkTestCrocking_Result = new DatabaseObject.ResultModel.FabricCrkShrkTestCrocking_Result()
            {
                ID = 999999,
                ScaleIDs = new List<string>() { "a", "b" },
                CrockingTestOption = 1, // 0 會隱藏明細欄位
                Crocking_Main = new DatabaseObject.ResultModel.FabricCrkShrkTestCrocking_Main
                {
                    Crocking = "Fail",
                    NonCrocking = true,
                    POID = "21051318BB",
                    SEQ = "01 01",
                    ColorID = "1",
                    ArriveQty = 2,
                    WhseArrival = System.DateTime.Now,
                    ExportID = "3",
                    Supp = "4",
                    CrockingDate = System.DateTime.Now.AddDays(-1),
                    StyleID = "5",
                    SCIRefno = "6",
                    Name = "7",
                    BrandID = "8",
                    Refno = "9",
                    DescDetail = "10",
                    CrockingRemark = "11",
                    CrockingEncdoe = true
                },
                Crocking_Detail = new List<DatabaseObject.ResultModel.FabricCrkShrkTestCrocking_Detail>(),
            };

            FabricCrkShrkTestCrocking_Detail test1 = new FabricCrkShrkTestCrocking_Detail()
            {
                Roll = "1",
                Dyelot = "2",
                Result = "Fail",
                DryScale = "a",
                ResultDry = "Fail",
                DryScale_Weft = "b",
                ResultDry_Weft = "Pass",
                WetScale = "a",
                ResultWet = "Fail",
                WetScale_Weft = "b",
                ResultWet_Weft = "Fail",
                Inspdate = System.DateTime.Now,
                Inspector = "3",
                Name = "3",
                Remark = "3",
                LastUpdate = "3"
            };
            fabricCrkShrkTestCrocking_Result.Crocking_Detail.Add(test1);

            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            ViewBag.ScaleIDsList = new SetListItem().ItemListBinding(fabricCrkShrkTestCrocking_Result.ScaleIDs);

            return View(fabricCrkShrkTestCrocking_Result);
        }

        public ActionResult HeatTest(long ID)
        {
            return View();
        }


        public ActionResult WashTest(long ID)
        {
            return View();
        }
    }
}