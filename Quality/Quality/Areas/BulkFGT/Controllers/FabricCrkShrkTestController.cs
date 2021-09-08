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
                CrockingTestOption = 0, // 0 會隱藏明細欄位
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
        public ActionResult AddCrockingDetailRow(int lastNO)
        {
            lastNO = lastNO + 1;

            // 假資料，因為清單來源目前沒有            
            string testTemp = "<option value='a'>a</option ><option value='b'>b</option>";

             string html = string.Empty;
            html += $"<tr idx='{lastNO}' class='row-content' style='vertical-align:middle;text-align:center;'>";
            html += "<td><div class='input-group'>";
            html += $"<input id='Crocking_Detail_{lastNO}__Roll' name='Crocking_Detail[{lastNO}].Roll' class='inputRollSelectItem' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</div></td>";
            html += "<td><div class='input-group'>";
            html += $"<input id='Crocking_Detail_{lastNO}__Dyelot' name='Crocking_Detail[{lastNO}].Dyelot' class='inputRollSelectItem' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</div></td>";
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
            html += $"<input class='date-picker' data-val='true' data-val-date='欄位 Inspdate 必須是日期。' id='Crocking_Detail_{lastNO}__Inspdate' name='Crocking_Detail[{lastNO}].Inspdate' type='text' value=''>";
            html += "</td>";
            html += "<td><div class='input-group'>";
            html += $"<input id='Crocking_Detail_{lastNO}__Inspector' name='Crocking_Detail[{lastNO}].Inspector' type='text' value=''>";
            html += $"<input id='btnDetailInspectorSelectItem' type='button' class='site-btn btn-blue' style='margin:0;border:0;' value='...'>";
            html += "</div></td>";
            html += "<td>";
            html += $"<input id='Crocking_Detail_{lastNO}__Name' name='Crocking_Detail[{lastNO}].Name' type='text' value=''>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Crocking_Detail_{lastNO}__Remark' name='Crocking_Detail[{lastNO}].Remark' type='text' value=''>";
            html += "</td>";
            html += "<td>";// last date
            html += "</td>";
            html += "<td><img class='detailDelete' src='/Image/Icon/Delete.png' width='30' style='min-width:30px'></td>";
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

            FabricCrkShrkTestCrocking_Detail test2 = new FabricCrkShrkTestCrocking_Detail()
            {
                Roll = "2",
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
            fabricCrkShrkTestCrocking_Result.Crocking_Detail.Add(test2);

            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            ViewBag.ScaleIDsList = new SetListItem().ItemListBinding(fabricCrkShrkTestCrocking_Result.ScaleIDs);

            return View(fabricCrkShrkTestCrocking_Result);
        }

        public ActionResult HeatTest(long ID)
        {
            DatabaseObject.ResultModel.FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result = new FabricCrkShrkTestHeat_Result()
            {
                ID = ID,
                Heat_Main = new FabricCrkShrkTestHeat_Main
                {
                    POID = "21051318BB",
                    SEQ = "01 02",
                    ColorID = "a",
                    ArriveQty = 1,
                    WhseArrival = System.DateTime.Now,
                    ExportID = "b",
                    Supp = "c",
                    Heat = "d",
                    HeatDate = System.DateTime.Now,
                    StyleID = "e",
                    SCIRefno = "f",
                    Name = "a",
                    BrandID = "b",
                    Refno = "c",
                    NonHeat = true,
                    DescDetail = "e",
                    HeatRemark = "f",
                    HeatEncode = false
                },
                Heat_Detail = new List<FabricCrkShrkTestHeat_Detail>()
            };

            FabricCrkShrkTestHeat_Detail test1 = new FabricCrkShrkTestHeat_Detail()
            {
                Roll = "a",
                Dyelot = "b",
                HorizontalOriginal = 1,
                VerticalOriginal = 2,
                Result = "Pass",
                HorizontalTest1 = 3,
                HorizontalTest2 = 4,
                HorizontalTest3 = 5,
                HorizontalAverage = 6,
                HorizontalRate = 7,
                VerticalTest1 = 8,
                VerticalTest2 = 9,
                VerticalTest3 = 10,
                VerticalAverage = 11,
                VerticalRate = 12,
                Inspdate = System.DateTime.Now,
                Inspector = "a",
                Name = "c",
                Remark = "d",
                LastUpdate = "e",
            };

            fabricCrkShrkTestHeat_Result.Heat_Detail.Add(test1);

            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            return View(fabricCrkShrkTestHeat_Result);
        }

        public ActionResult AddHeatDetailRow(int lastNO)
        {
            lastNO = lastNO + 1;
            string html = string.Empty;
            html += $"<tr idx='{lastNO}' class='row-content' style='vertical-align: middle; text-align: center;'>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputRollSelectItem' id='Heat_Detail_{lastNO}__Roll' name='Heat_Detail[{lastNO}].Roll' type='text' value='' >";
            html += "<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...' >";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputRollSelectItem' id='Heat_Detail_{lastNO}__Dyelot' name='Heat_Detail[{lastNO}].Dyelot' type='text' value='' >";
            html += "<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...' >";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalOriginal 必須是數字。' data-val-required='HorizontalOriginal 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalOriginal' name='Heat_Detail[{lastNO}].HorizontalOriginal' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalOriginal 必須是數字。' data-val-required='VerticalOriginal 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalOriginal' name='Heat_Detail[{lastNO}].VerticalOriginal' step='0.01' type='number' value='0' >";
            html += "</td>";

            html += "<td>";
            html += "<select id='Heat_Detail_0__Result' class='resultList' name='Model.Heat_Detail[0].Result' style='width:157px;color:blue' onchange='changeResultColor(this)'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest1 必須是數字。' data-val-required='HorizontalTest1 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalTest1' name='Heat_Detail[{lastNO}].HorizontalTest1' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest2 必須是數字。' data-val-required='HorizontalTest2 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalTest2' name='Heat_Detail[{lastNO}].HorizontalTest2' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest3 必須是數字。' data-val-required='HorizontalTest3 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalTest3' name='Heat_Detail[{lastNO}].HorizontalTest3' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalAverage 必須是數字。' data-val-required='HorizontalAverage 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalAverage' name='Heat_Detail[{lastNO}].HorizontalAverage' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalRate 必須是數字。' data-val-required='HorizontalRate 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalRate' name='Heat_Detail[{lastNO}].HorizontalRate' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest1 必須是數字。' data-val-required='VerticalTest1 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalTest1' name='Heat_Detail[{lastNO}].VerticalTest1' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest2 必須是數字。' data-val-required='VerticalTest2 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalTest2' name='Heat_Detail[{lastNO}].VerticalTest2' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest3 必須是數字。' data-val-required='VerticalTest3 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalTest3' name='Heat_Detail[{lastNO}].VerticalTest3' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalAverage 必須是數字。' data-val-required='VerticalAverage 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalAverage' name='Heat_Detail[{lastNO}].VerticalAverage' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalRate 必須是數字。' data-val-required='VerticalRate 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalRate' name='Heat_Detail[{lastNO}].VerticalRate' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='date-picker' data-val='true' data-val-date='欄位 Inspdate 必須是日期。' id='Heat_Detail_{lastNO}__Inspdate' name='Heat_Detail[{lastNO}].Inspdate' type='text' value='' >";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputInspectorSelectItem' id='Heat_Detail_{lastNO}__Inspector' name='Heat_Detail[{lastNO}].Inspector' type='text' value='' >";
            html += "<input id='btnDetailInspectorSelectItem' type='button' class='site-btn btn-blue btnInspectorSelectItem' style='margin:0;border:0;' value='...' >";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Heat_Detail_{lastNO}__Name' name='Heat_Detail[{lastNO}].Name' readonly='readonly' type='text' value='' >";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Heat_Detail_{lastNO}__Remark' name='Heat_Detail[{lastNO}].Remark' readonly='readonly' type='text' value='' >";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Heat_Detail_{lastNO}__LastUpdate' name='Heat_Detail[{lastNO}].LastUpdate' readonly='readonly' type='text' value='' >";
            html += "</td>";
            html += "<td>";
            html += "<img class='detailDelete' src='/Image/Icon/Delete.png' style='min-width:30px'>";
            html += "</td>";
            html += "</tr>";
            return Content(html);
        }

        [HttpPost]
        public ActionResult HeatTest(DatabaseObject.ResultModel.FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result)
        {
            fabricCrkShrkTestHeat_Result = new FabricCrkShrkTestHeat_Result()
            {
                ID = ID,
                Heat_Main = new FabricCrkShrkTestHeat_Main
                {
                    POID = "21051318BB",
                    SEQ = "01 02",
                    ColorID = "a",
                    ArriveQty = 1,
                    WhseArrival = System.DateTime.Now,
                    ExportID = "b",
                    Supp = "c",
                    Heat = "d",
                    HeatDate = System.DateTime.Now,
                    StyleID = "e",
                    SCIRefno = "f",
                    Name = "a",
                    BrandID = "b",
                    Refno = "c",
                    NonHeat = true,
                    DescDetail = "e",
                    HeatRemark = "f",
                    HeatEncode = false
                },
                Heat_Detail = new List<FabricCrkShrkTestHeat_Detail>()
            };

            FabricCrkShrkTestHeat_Detail test1 = new FabricCrkShrkTestHeat_Detail()
            {
                Roll = "a",
                Dyelot = "b",
                HorizontalOriginal = 1,
                VerticalOriginal = 2,
                Result = "Pass",
                HorizontalTest1 = 3,
                HorizontalTest2 = 4,
                HorizontalTest3 = 5,
                HorizontalAverage = 6,
                HorizontalRate = 7,
                VerticalTest1 = 8,
                VerticalTest2 = 9,
                VerticalTest3 = 10,
                VerticalAverage = 11,
                VerticalRate = 12,
                Inspdate = System.DateTime.Now,
                Inspector = "a",
                Name = "c",
                Remark = "d",
                LastUpdate = "e",
            };

            fabricCrkShrkTestHeat_Result.Heat_Detail.Add(test1);

            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            return View(fabricCrkShrkTestHeat_Result);
        }

        public ActionResult WashTest(long ID)
        {
            DatabaseObject.ResultModel.FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result = new FabricCrkShrkTestWash_Result()
            {
                ID = 9997,
                Wash_Main = new FabricCrkShrkTestWash_Main
                {
                    POID = "21051318BB",
                    SEQ = "01 02",
                    ColorID = "a",
                    ArriveQty = 1,
                    WhseArrival = System.DateTime.Now,
                    ExportID = "b",
                    Supp = "c",
                    Wash = "d",
                    WashDate = System.DateTime.Now,
                    StyleID = "e",
                    SCIRefno = "f",
                    Name = "a",
                    BrandID = "b",
                    Refno = "c",
                    NonWash = true,
                    SkewnessOptionID = "Option2",
                    DescDetail = "d",
                    WashRemark = "e",
                    WashEncode = false
                },
                Wash_Detail = new List<FabricCrkShrkTestWash_Detail>()
            };

            FabricCrkShrkTestWash_Detail test1 = new FabricCrkShrkTestWash_Detail()
            {
                Roll = "a",
                Dyelot = "b",
                HorizontalOriginal = 1,
                VerticalOriginal = 2,
                Result = "Pass",
                HorizontalTest1 = 3,
                HorizontalTest2 = 4,
                HorizontalTest3 = 5,
                HorizontalAverage = 0,
                HorizontalRate = 0,
                VerticalTest1 = 1,
                VerticalTest2 = 2,
                VerticalTest3 = 3,
                VerticalAverage = 0,
                VerticalRate = 0,
                SkewnessTest1 = 4,
                SkewnessTest2 = 5,
                SkewnessTest3 = 6,
                SkewnessTest4 = 7,
                SkewnessRate = 8,
                Inspdate = System.DateTime.Now,
                Inspector = "A",
                Name = "",
                Remark = "",
                LastUpdate = "bbb",
            };

            fabricCrkShrkTestWash_Result.Wash_Detail.Add(test1);

            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };

            Dictionary<string, string> skewnessOption = new Dictionary<string, string>() {
                { "Option1", "1" },
                { "Option2", "2" },
                { "Option3", "3" }
            };

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<SelectListItem> result_itemList = new List<SelectListItem>();
            foreach (var item in skewnessOption)
            {
                SelectListItem i = new SelectListItem()
                {
                    Text = item.Key,
                    Value = item.Value,
                };
                result_itemList.Add(i);
            }

            ViewBag.SkewnessOptionList = result_itemList;
            return View(fabricCrkShrkTestWash_Result);
        }

        public ActionResult AddWashDetailRow(int lastNO)
        {
            lastNO = lastNO + 1;
            string html = string.Empty;
            html += $"<tr idx='{lastNO}' class='row-content' style='vertical-align: middle; text-align: center;'>";
            html += "<td>";
            html += "<div class='input -group'>";
            html += $"<input class='inputRollSelectItem' id='Wash_Detail_{lastNO}__Roll' name='Wash_Detail[{lastNO}].Roll' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputRollSelectItem' id='Wash_Detail_{lastNO}__Dyelot' name='Wash_Detail[{lastNO}].Dyelot' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalOriginal 必須是數字。' data-val-required='HorizontalOriginal 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalOriginal' name='Wash_Detail[{lastNO}].HorizontalOriginal' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalOriginal 必須是數字。' data-val-required='VerticalOriginal 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalOriginal' name='Wash_Detail[{lastNO}].VerticalOriginal' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Wash_Main_{lastNO}__Result' class='resultList' name='Model.Wash_Detail[{lastNO}].Result' style='width:157px;color:blue' onchange='changeResultColor(this)'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest1 必須是數字。' data-val-required='HorizontalTest1 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalTest1' name='Wash_Detail[{lastNO}].HorizontalTest1' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest2 必須是數字。' data-val-required='HorizontalTest2 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalTest2' name='Wash_Detail[{lastNO}].HorizontalTest2' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest3 必須是數字。' data-val-required='HorizontalTest3 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalTest3' name='Wash_Detail[{lastNO}].HorizontalTest3' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalAverage 必須是數字。' data-val-required='HorizontalAverage 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalAverage' name='Wash_Detail[{lastNO}].HorizontalAverage' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalRate 必須是數字。' data-val-required='HorizontalRate 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalRate' name='Wash_Detail[{lastNO}].HorizontalRate' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest1 必須是數字。' data-val-required='VerticalTest1 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalTest1' name='Wash_Detail[{lastNO}].VerticalTest1' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest2 必須是數字。' data-val-required='VerticalTest2 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalTest2' name='Wash_Detail[{lastNO}].VerticalTest2' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest3 必須是數字。' data-val-required='VerticalTest3 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalTest3' name='Wash_Detail[{lastNO}].VerticalTest3' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalAverage 必須是數字。' data-val-required='VerticalAverage 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalAverage' name='Wash_Detail[{lastNO}].VerticalAverage' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalRate 必須是數字。' data-val-required='VerticalRate 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalRate' name='Wash_Detail[{lastNO}].VerticalRate' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest1 必須是數字。' data-val-required='SkewnessTest1 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest1' name='Wash_Detail[{lastNO}].SkewnessTest1' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest2 必須是數字。' data-val-required='SkewnessTest2 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest2' name='Wash_Detail[{lastNO}].SkewnessTest2' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td style=''>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest3 必須是數字。' data-val-required='SkewnessTest3 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest3' name='Wash_Detail[{lastNO}].SkewnessTest3' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td style=''>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest4 必須是數字。' data-val-required='SkewnessTest4 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest4' name='Wash_Detail[{lastNO}].SkewnessTest4' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 SkewnessRate 必須是數字。' data-val-required='SkewnessRate 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessRate' name='Wash_Detail[{lastNO}].SkewnessRate' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='date-picker' data-val='true' data-val-date='欄位 Inspdate 必須是日期。' id='Wash_Detail_{lastNO}__Inspdate' name='Wash_Detail[{lastNO}].Inspdate' type='text' value=''>";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputInspectorSelectItem' id='Wash_Detail_{lastNO}__Inspector' name='Wash_Detail[{lastNO}].Inspector' type='text' value=''>";
            html += $"<input id='btnDetailInspectorSelectItem' type='button' class='site-btn btn-blue btnInspectorSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Wash_Detail_{lastNO}__Name' name='Wash_Detail[{lastNO}].Name' readonly='readonly' type='text' value=''>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Wash_Detail_{lastNO}__Remark' name='Wash_Detail[{lastNO}].Remark' readonly='readonly' type='text' value=''>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Wash_Detail_{lastNO}__LastUpdate' name='Wash_Detail[{lastNO}].LastUpdate' readonly='readonly' type='text' value=''>";
            html += "</td>";
            html += "<td>";
            html += "<img class='detailDelete' src='/Image/Icon/Delete.png' style='min-width:30px'>";
            html += "</td>";
            html += "</tr>";
            return Content(html);
        }

        [HttpPost]
        public ActionResult WashTest(DatabaseObject.ResultModel.FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result)
        {
            fabricCrkShrkTestWash_Result = new FabricCrkShrkTestWash_Result()
            {
                ID = 9997,
                Wash_Main = new FabricCrkShrkTestWash_Main
                {
                    POID = "21051318BB",
                    SEQ = "01 02",
                    ColorID = "a",
                    ArriveQty = 1,
                    WhseArrival = System.DateTime.Now,
                    ExportID = "b",
                    Supp = "c",
                    Wash = "d",
                    WashDate = System.DateTime.Now,
                    StyleID = "e",
                    SCIRefno = "f",
                    Name = "a",
                    BrandID = "b",
                    Refno = "c",
                    NonWash = true,
                    SkewnessOptionID = "Option2",
                    DescDetail = "d",
                    WashRemark = "e",
                    WashEncode = false
                },
                Wash_Detail = new List<FabricCrkShrkTestWash_Detail>()
            };

            FabricCrkShrkTestWash_Detail test1 = new FabricCrkShrkTestWash_Detail()
            {
                Roll = "a",
                Dyelot = "b",
                HorizontalOriginal = 1,
                VerticalOriginal = 2,
                Result = "Pass",
                HorizontalTest1 = 3,
                HorizontalTest2 = 4,
                HorizontalTest3 = 5,
                HorizontalAverage = 0,
                HorizontalRate = 0,
                VerticalTest1 = 1,
                VerticalTest2 = 2,
                VerticalTest3 = 3,
                VerticalAverage = 0,
                VerticalRate = 0,
                SkewnessTest1 = 4,
                SkewnessTest2 = 5,
                SkewnessTest3 = 6,
                SkewnessTest4 = 7,
                SkewnessRate = 8,
                Inspdate = System.DateTime.Now,
                Inspector = "A",
                Name = "",
                Remark = "",
                LastUpdate = "bbb",
            };

            fabricCrkShrkTestWash_Result.Wash_Detail.Add(test1);

            List<string> resultType = new List<string>() {
                 "Pass","Fail"
            };

            Dictionary<string, string> skewnessOption = new Dictionary<string, string>() {
                { "Option1", "1" },
                { "Option2", "2" },
                { "Option3", "3" }
            };

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);

            List<SelectListItem> result_itemList = new List<SelectListItem>();
            foreach (var item in skewnessOption)
            {
                SelectListItem i = new SelectListItem()
                {
                    Text = item.Key,
                    Value = item.Value,
                };
                result_itemList.Add(i);
            }

            ViewBag.SkewnessOptionList = result_itemList;
            return View(fabricCrkShrkTestWash_Result);
        }
    }
}