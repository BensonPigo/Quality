using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
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
    public class FabricCrkShrkTestController : BaseController
    {
        private IFabricCrkShrkTest_Service _FabricCrkShrkTest_Service;
        private List<string> resultType = new List<string>() { "Pass", "Fail" };
        private Dictionary<string, string> skewnessOption = new Dictionary<string, string>() { { "Option1", "1" }, { "Option2", "2" }, { "Option3", "3" } };
        public FabricCrkShrkTestController()
        {
            _FabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.FabricCrkShrkTest,,";
        }

        // GET: BulkFGT/FabricCrkShrkTest
        public ActionResult Index()
        {
            FabricCrkShrkTest_Result fabricCrkShrkTest_Result = new FabricCrkShrkTest_Result()
            {
                Main = new FabricCrkShrkTest_Main(),
                Details = new List<FabricCrkShrkTest_Detail>()
            };

            return View(fabricCrkShrkTest_Result);
        }

        [HttpPost]
        public ActionResult Index(string POID)
        {
            ViewBag.POID = POID;
            FabricCrkShrkTest_Result fabricCrkShrkTest_Result = _FabricCrkShrkTest_Service.GetFabricCrkShrkTest_Result(POID);
            if (!fabricCrkShrkTest_Result.Result)
            {
                fabricCrkShrkTest_Result = new FabricCrkShrkTest_Result() 
                {
                    Main = new FabricCrkShrkTest_Main(),
                    Details = new List<FabricCrkShrkTest_Detail>(),
                    Result = false,
                    ErrorMessage = $"msg.WithInfo('{fabricCrkShrkTest_Result.ErrorMessage.Replace("\r\n", "<br />") }');",
                };
            }
            UpdateModel(fabricCrkShrkTest_Result);
            return View(fabricCrkShrkTest_Result);
        }

        [HttpPost]
        public ActionResult SaveIndex(FabricCrkShrkTest_Main main, List<FabricCrkShrkTest_Detail> detail)
        {
            FabricCrkShrkTest_Result result = new FabricCrkShrkTest_Result()
                                              {
                                                  Main = main,
                                                  Details = detail,
                                              };

            BaseResult saveResult = _FabricCrkShrkTest_Service.SaveFabricCrkShrkTestMain(result);
            return Json(saveResult);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        [HttpGet]
        public ActionResult IndexBack(string POID)
        {
            ViewBag.POID = POID;
            FabricCrkShrkTest_Result fabricCrkShrkTest_Result = _FabricCrkShrkTest_Service.GetFabricCrkShrkTest_Result(POID);
            return View("Index", fabricCrkShrkTest_Result);
        }

        #region Crocking
        public ActionResult CrockingTest(long ID)
        {
            FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result = _FabricCrkShrkTest_Service.GetFabricCrkShrkTestCrocking_Result(ID);
            if (!fabricCrkShrkTestCrocking_Result.Result)
            {
                fabricCrkShrkTestCrocking_Result = new FabricCrkShrkTestCrocking_Result()
                {
                    ScaleIDs = new List<string>(),
                    Crocking_Main = new FabricCrkShrkTestCrocking_Main(),
                    Crocking_Detail = new List<FabricCrkShrkTestCrocking_Detail>(),
                    ErrorMessage = $"msg.WithInfo('{fabricCrkShrkTestCrocking_Result.ErrorMessage.Replace("\r\n", "<br />") }');",
                };
            }

            if (TempData["Model"] != null)
            {
                FabricCrkShrkTestCrocking_Result saveResult = (FabricCrkShrkTestCrocking_Result)TempData["Model"];
                fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingRemark = saveResult.Crocking_Main.CrockingRemark;
                fabricCrkShrkTestCrocking_Result.Crocking_Detail = saveResult.Crocking_Detail;
                fabricCrkShrkTestCrocking_Result.Result = saveResult.Result;
                fabricCrkShrkTestCrocking_Result.ErrorMessage = $"msg.WithInfo('{saveResult.ErrorMessage.ToString().Replace("\r\n", "<br />") }');EditMode = true;";
            }

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            ViewBag.ScaleIDsList = new SetListItem().ItemListBinding(fabricCrkShrkTestCrocking_Result.ScaleIDs);
            ViewBag.FactoryID = this.FactoryID;
            return View(fabricCrkShrkTestCrocking_Result);
        }

        [HttpPost]
        public ActionResult AddCrockingDetailRow(int lastNO)
        {
            List<string> scaleIDs = _FabricCrkShrkTest_Service.GetScaleIDs();
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
            html += $"<input id='Crocking_Detail_{lastNO}__Result' name='Crocking_Detail[{lastNO}].Result' readonly='readonly' class='blue' type='text' value='Pass'>";
            html += "</td>";
            html += "<td style='color:blue'>";
            html += $"<select id='Crocking_Detail_{lastNO}__DryScale' name='Crocking_Detail[{lastNO}].DryScale' style='width:157px'>";
            foreach (string val in scaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select class='result{lastNO}' id='Crocking_Detail_{lastNO}__ResultDry' name='Crocking_Detail[{lastNO}].ResultDry' onchange='changeResultColor(this)' style='width:157px'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Crocking_Detail_{lastNO}__DryScale_Weft' name='Crocking_Detail[{lastNO}].DryScale_Weft' style='width:157px'>";
            foreach (string val in scaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select class='result{lastNO}' id='Crocking_Detail_{lastNO}__ResultDry_Weft' name='Crocking_Detail[{lastNO}].ResultDry_Weft' onchange='changeResultColor(this)' style='width:157px'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Crocking_Detail_{lastNO}__WetScale' name='Crocking_Detail[{lastNO}].WetScale' style='width:157px'>";
            foreach (string val in scaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select class='result{lastNO}' id='Crocking_Detail_{lastNO}__ResultWet' name='Crocking_Detail[{lastNO}].ResultWet' onchange='changeResultColor(this)' style='width:157px'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Crocking_Detail_{lastNO}__WetScale_Weft' name='Crocking_Detail[{lastNO}].WetScale_Weft' style='width:157px'>";
            foreach (string val in scaleIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<select class='result{lastNO}' id='Crocking_Detail_{lastNO}__ResultWet_Weft' name='Crocking_Detail[{lastNO}].ResultWet_Weft' onchange='changeResultColor(this)' style='width:157px'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='date-picker' data-val='true' data-val-date='欄位 Inspdate 必須是日期。' id='Crocking_Detail_{lastNO}__Inspdate' name='Crocking_Detail[{lastNO}].Inspdate' type='text' value=''>";
            html += "</td>";
            html += "<td><div class='input-group'>";
            html += $"<input id='Crocking_Detail_{lastNO}__Inspector' name='Crocking_Detail[{lastNO}].Inspector' type='text' value='' class='inputInspectorSelectItem'>";
            html += $"<input id='btnDetailInspectorSelectItem' type='button' class='site-btn btn-blue btnInspectorSelectItem' style='margin:0;border:0;' value='...'>";
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
        public ActionResult CrockingTestSave(FabricCrkShrkTestCrocking_Result Result)
        {
            if (Result.Crocking_Detail == null)
            {
                Result.Crocking_Detail = new List<FabricCrkShrkTestCrocking_Detail>();
            }

            BaseResult saveResult = _FabricCrkShrkTest_Service.SaveFabricCrkShrkTestCrockingDetail(Result, this.UserID);
            if (saveResult.Result)
            {
                return RedirectToAction("CrockingTest", new { ID = Result.ID });
            }
            Result.Result = saveResult.Result;
            Result.ErrorMessage = saveResult.ErrorMessage;
            TempData["Model"] = Result;
            return RedirectToAction("CrockingTest", new { ID = Result.ID });
        }

        [HttpPost]
        public JsonResult Encode_Crocking(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.EncodeFabricCrkShrkTestCrockingDetail(ID, this.UserID, out string testResult);
            return Json(new { result.Result, result.ErrorMessage, testResult });
        }

        [HttpPost]
        public JsonResult Amend_Crocking(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.AmendFabricCrkShrkTestCrockingDetail(ID);
            return Json(new { result.Result, result.ErrorMessage });
        }

        [HttpPost]
        public JsonResult FailMail_Crocking(long ID, string TO, string CC)
        {
            SendMail_Result result = _FabricCrkShrkTest_Service.SendCrockingFailResultMail(TO, CC, ID, false);
            return Json(result);
        }

        [HttpPost]
        public JsonResult Report_Crocking(long ID, bool IsToPDF)
        {
            BaseResult result;
            string FileName;
            if (IsToPDF)
            {
                result = _FabricCrkShrkTest_Service.ToPdfFabricCrkShrkTestCrockingDetail(ID, out FileName, false);
            }
            else
            {
                result = _FabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestCrockingDetail(ID, out FileName, false);
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        #endregion

        #region Heat
        public ActionResult HeatTest(long ID)
        {
            FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result = _FabricCrkShrkTest_Service.GetFabricCrkShrkTestHeat_Result(ID);
            if (!fabricCrkShrkTestHeat_Result.Result)
            {
                fabricCrkShrkTestHeat_Result = new FabricCrkShrkTestHeat_Result()
                {
                    Heat_Main = new FabricCrkShrkTestHeat_Main(),
                    Heat_Detail = new List<FabricCrkShrkTestHeat_Detail>(),
                    ErrorMessage = $"msg.WithInfo('{fabricCrkShrkTestHeat_Result.ErrorMessage.Replace("\r\n", "<br />") }');",
                };
            }

            if (TempData["Model"] != null)
            {
                FabricCrkShrkTestHeat_Result saveResult = (FabricCrkShrkTestHeat_Result)TempData["Model"];
                fabricCrkShrkTestHeat_Result.Heat_Main.HeatRemark = saveResult.Heat_Main.HeatRemark;
                fabricCrkShrkTestHeat_Result.Heat_Detail = saveResult.Heat_Detail;
                fabricCrkShrkTestHeat_Result.Result = saveResult.Result;
                fabricCrkShrkTestHeat_Result.ErrorMessage = $"msg.WithInfo('{saveResult.ErrorMessage.ToString().Replace("\r\n", "<br />") }');EditMode = true;";
            }

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            ViewBag.FactoryID = this.FactoryID;
            return View(fabricCrkShrkTestHeat_Result);
        }

        public ActionResult AddHeatDetailRow(int lastNO)
        {
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
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalOriginal 必須是數字。' data-val-required='HorizontalOriginal 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalOriginal' name='Heat_Detail[{lastNO}].HorizontalOriginal' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalOriginal 必須是數字。' data-val-required='VerticalOriginal 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalOriginal' name='Heat_Detail[{lastNO}].VerticalOriginal' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";

            html += "<td>";
            html += $"<select id='Heat_Detail_{lastNO}__Result' name='Heat_Detail[{lastNO}].Result' style='width:157px;color:blue' onchange='changeResultColor(this)'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest1 必須是數字。' data-val-required='HorizontalTest1 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalTest1' name='Heat_Detail[{lastNO}].HorizontalTest1' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest2 必須是數字。' data-val-required='HorizontalTest2 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalTest2' name='Heat_Detail[{lastNO}].HorizontalTest2' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest3 必須是數字。' data-val-required='HorizontalTest3 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalTest3' name='Heat_Detail[{lastNO}].HorizontalTest3' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalAverage 必須是數字。' data-val-required='HorizontalAverage 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalAverage' name='Heat_Detail[{lastNO}].HorizontalAverage' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalRate 必須是數字。' data-val-required='HorizontalRate 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalRate' name='Heat_Detail[{lastNO}].HorizontalRate' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest1 必須是數字。' data-val-required='VerticalTest1 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalTest1' name='Heat_Detail[{lastNO}].VerticalTest1' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest2 必須是數字。' data-val-required='VerticalTest2 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalTest2' name='Heat_Detail[{lastNO}].VerticalTest2' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest3 必須是數字。' data-val-required='VerticalTest3 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalTest3' name='Heat_Detail[{lastNO}].VerticalTest3' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
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
        public ActionResult HeatTestSave(FabricCrkShrkTestHeat_Result Result)
        {
            if (Result.Heat_Detail == null)
            {
                Result.Heat_Detail = new List<FabricCrkShrkTestHeat_Detail>();
            }

            BaseResult saveResult = _FabricCrkShrkTest_Service.SaveFabricCrkShrkTestHeatDetail(Result, this.UserID);

            if (saveResult.Result)
            {
                return RedirectToAction("HeatTest", new { ID = Result.ID });
            }
            Result.Result = saveResult.Result;
            Result.ErrorMessage = saveResult.ErrorMessage;
            TempData["Model"] = Result;
            return RedirectToAction("HeatTest", new { ID = Result.ID });
        }

        [HttpPost]
        public JsonResult Encode_Heat(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.EncodeFabricCrkShrkTestHeatDetail(ID, this.UserID, out string testResult);
            return Json(new { result.Result, result.ErrorMessage, testResult });
        }

        [HttpPost]
        public JsonResult Amend_Heat(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.AmendFabricCrkShrkTestHeatDetail(ID);
            return Json(new { result.Result, result.ErrorMessage });
        }

        [HttpPost]
        public JsonResult FailMail_Heat(long ID, string TO, string CC)
        {
            SendMail_Result result = _FabricCrkShrkTest_Service.SendHeatFailResultMail(TO, CC, ID, false);
            return Json(result);
        }

        [HttpPost]
        public JsonResult Report_Heat(long ID)
        {
            BaseResult result;
            result = _FabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestHeatDetail(ID, out string FileName, false);
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        #endregion

        #region Wash
        public ActionResult WashTest(long ID)
        {
            FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result = _FabricCrkShrkTest_Service.GetFabricCrkShrkTestWash_Result(ID);
            if (!fabricCrkShrkTestWash_Result.Result)
            {
                fabricCrkShrkTestWash_Result = new FabricCrkShrkTestWash_Result()
                {
                    Wash_Main = new FabricCrkShrkTestWash_Main(),
                    Wash_Detail = new List<FabricCrkShrkTestWash_Detail>(),
                    ErrorMessage = $"msg.WithInfo('{fabricCrkShrkTestWash_Result.ErrorMessage.Replace("\r\n", "<br />") }');",
                };
            }

            if (TempData["Model"] != null)
            {
                FabricCrkShrkTestWash_Result saveResult = (FabricCrkShrkTestWash_Result)TempData["Model"];
                fabricCrkShrkTestWash_Result.Wash_Main.SkewnessOptionID = saveResult.Wash_Main.SkewnessOptionID;
                fabricCrkShrkTestWash_Result.Wash_Main.WashRemark = saveResult.Wash_Main.WashRemark;
                fabricCrkShrkTestWash_Result.Wash_Detail = saveResult.Wash_Detail;
                fabricCrkShrkTestWash_Result.Result = saveResult.Result;
                fabricCrkShrkTestWash_Result.ErrorMessage = $"msg.WithInfo('{saveResult.ErrorMessage.ToString().Replace("\r\n", "<br />") }');EditMode = true;";
            }

            List<SelectListItem> skewnessOptionList = new SetListItem().ItemListBinding(skewnessOption);
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            ViewBag.SkewnessOptionList = skewnessOptionList;
            ViewBag.FactoryID = this.FactoryID;
            return View(fabricCrkShrkTestWash_Result);
        }

        [HttpPost]
        public ActionResult AddWashDetailRow(int lastNO)
        {
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
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalOriginal 必須是數字。' data-val-required='HorizontalOriginal 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalOriginal' name='Wash_Detail[{lastNO}].HorizontalOriginal' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalOriginal 必須是數字。' data-val-required='VerticalOriginal 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalOriginal' name='Wash_Detail[{lastNO}].VerticalOriginal' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Wash_Main_{lastNO}__Result' name='Wash_Detail[{lastNO}].Result' style='width:157px;color:blue' onchange='changeResultColor(this)'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest1 必須是數字。' data-val-required='HorizontalTest1 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalTest1' name='Wash_Detail[{lastNO}].HorizontalTest1' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest2 必須是數字。' data-val-required='HorizontalTest2 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalTest2' name='Wash_Detail[{lastNO}].HorizontalTest2' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest3 必須是數字。' data-val-required='HorizontalTest3 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalTest3' name='Wash_Detail[{lastNO}].HorizontalTest3' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalAverage 必須是數字。' data-val-required='HorizontalAverage 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalAverage' name='Wash_Detail[{lastNO}].HorizontalAverage' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalRate 必須是數字。' data-val-required='HorizontalRate 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalRate' name='Wash_Detail[{lastNO}].HorizontalRate' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest1 必須是數字。' data-val-required='VerticalTest1 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalTest1' name='Wash_Detail[{lastNO}].VerticalTest1' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest2 必須是數字。' data-val-required='VerticalTest2 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalTest2' name='Wash_Detail[{lastNO}].VerticalTest2' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest3 必須是數字。' data-val-required='VerticalTest3 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalTest3' name='Wash_Detail[{lastNO}].VerticalTest3' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalAverage 必須是數字。' data-val-required='VerticalAverage 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalAverage' name='Wash_Detail[{lastNO}].VerticalAverage' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalRate 必須是數字。' data-val-required='VerticalRate 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalRate' name='Wash_Detail[{lastNO}].VerticalRate' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest1 必須是數字。' data-val-required='SkewnessTest1 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest1' name='Wash_Detail[{lastNO}].SkewnessTest1' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest2 必須是數字。' data-val-required='SkewnessTest2 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest2' name='Wash_Detail[{lastNO}].SkewnessTest2' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td style=''>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest3 必須是數字。' data-val-required='SkewnessTest3 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest3' name='Wash_Detail[{lastNO}].SkewnessTest3' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td style=''>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest4 必須是數字。' data-val-required='SkewnessTest4 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest4' name='Wash_Detail[{lastNO}].SkewnessTest4' step='0.01' type='number' value='0' oninput='value=QtyCheck(value)'>";
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
        public ActionResult WashTestSave(FabricCrkShrkTestWash_Result Result)
        {
            if (Result.Wash_Detail == null)
            {
                Result.Wash_Detail = new List<FabricCrkShrkTestWash_Detail>();
            }

            BaseResult saveResult = _FabricCrkShrkTest_Service.SaveFabricCrkShrkTestWashDetail(Result, this.UserID);
            if (saveResult.Result)
            {
                return RedirectToAction("WashTest", new { ID = Result.ID });
            }

            Result.Result = saveResult.Result;
            Result.ErrorMessage = saveResult.ErrorMessage;
            TempData["Model"] = Result;
            return RedirectToAction("WashTest", new { ID = Result.ID });
        }

        [HttpPost]
        public JsonResult Encode_Wash(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.EncodeFabricCrkShrkTestWashDetail(ID, this.UserID, out string testResult);
            return Json(new { result.Result, result.ErrorMessage, testResult });
        }

        [HttpPost]
        public JsonResult Amend_Wash(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.AmendFabricCrkShrkTestWashDetail(ID);
            return Json(new { result.Result, result.ErrorMessage });
        }

        [HttpPost]
        public JsonResult FailMail_Wash(long ID, string TO, string CC)
        {
            SendMail_Result result = _FabricCrkShrkTest_Service.SendWashFailResultMail(TO, CC, ID, false);
            return Json(result);
        }

        [HttpPost]
        public JsonResult Report_Wash(long ID)
        {
            BaseResult result;
            result = _FabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestWashDetail(ID, out string FileName, false);
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }
        #endregion
    }
}