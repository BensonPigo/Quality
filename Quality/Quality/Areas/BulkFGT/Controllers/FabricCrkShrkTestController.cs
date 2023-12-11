using BusinessLogicLayer.Interface;
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
        [SessionAuthorizeAttribute]
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
        [SessionAuthorizeAttribute]
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
                    ErrorMessage = $@"msg.WithInfo('{(string.IsNullOrEmpty(fabricCrkShrkTest_Result.ErrorMessage) ? string.Empty : fabricCrkShrkTest_Result.ErrorMessage.Replace("'", string.Empty)) }');",
                };
            }
            UpdateModel(fabricCrkShrkTest_Result);
            return View(fabricCrkShrkTest_Result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
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
        [SessionAuthorizeAttribute]
        public ActionResult IndexBack(string POID)
        {
            ViewBag.POID = POID;
            FabricCrkShrkTest_Result fabricCrkShrkTest_Result = _FabricCrkShrkTest_Service.GetFabricCrkShrkTest_Result(POID);
            return View("Index", fabricCrkShrkTest_Result);
        }

        #region Crocking
        [SessionAuthorizeAttribute]
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
                    ErrorMessage = $@"msg.WithInfo('{ (string.IsNullOrEmpty(fabricCrkShrkTestCrocking_Result.ErrorMessage) ? string.Empty : fabricCrkShrkTestCrocking_Result.ErrorMessage.Replace("'", string.Empty)) }');",
                };
            }

            if (TempData["ModelCrockingTest"] != null)
            {
                FabricCrkShrkTestCrocking_Result saveResult = (FabricCrkShrkTestCrocking_Result)TempData["ModelCrockingTest"];
                fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingRemark = saveResult.Crocking_Main.CrockingRemark;
                fabricCrkShrkTestCrocking_Result.Crocking_Detail = saveResult.Crocking_Detail;
                fabricCrkShrkTestCrocking_Result.Result = saveResult.Result;
                fabricCrkShrkTestCrocking_Result.ErrorMessage = $@"msg.WithInfo('{(string.IsNullOrEmpty(saveResult.ErrorMessage) ? string.Empty : saveResult.ErrorMessage.Replace("'", string.Empty)) });EditMode = true;";
            }

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            ViewBag.ScaleIDsList = new SetListItem().ItemListBinding(fabricCrkShrkTestCrocking_Result.ScaleIDs);
            //ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View(fabricCrkShrkTestCrocking_Result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult AddCrockingDetailRow(int lastNO)
        {
            List<string> scaleIDs = _FabricCrkShrkTest_Service.GetScaleIDs();
            string html = string.Empty;
            html += $"<tr idx='{lastNO}' class='row-content' style='vertical-align:middle;text-align:center;'>";
            html += "<td><div class='input-group'>";
            html += $"<input id='Crocking_Detail_{lastNO}__Roll' name='Crocking_Detail[{lastNO}].Roll' class='inputRollSelectItem width6vw' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</div></td>";
            html += "<td><div class='input-group'>";
            html += $"<input id='Crocking_Detail_{lastNO}__Dyelot' name='Crocking_Detail[{lastNO}].Dyelot' class='inputRollSelectItem width8vw' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</div></td>";
            html += "<td>";
            html += $"<input id='Crocking_Detail_{lastNO}__Result' name='Crocking_Detail[{lastNO}].Result' readonly='readonly' class='blue width6vw' type='text' value='Pass'>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Crocking_Detail_{lastNO}__DryScale' name='Crocking_Detail[{lastNO}].DryScale' style='width:157px'>";
            foreach (string val in scaleIDs)
            {
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
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
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
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
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
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
                string selected = string.Empty;
                if (val == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val + "'>" + val + "</option>";
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
            html += $"<input class='date-picker width9vw' data-val='true' data-val-date='欄位 Inspdate 必須是日期。' id='Crocking_Detail_{lastNO}__Inspdate' name='Crocking_Detail[{lastNO}].Inspdate' type='text' value=''>";
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
        [SessionAuthorizeAttribute]
        public ActionResult CrockingTestSave(FabricCrkShrkTestCrocking_Result Result)
        {
            if (Result.Crocking_Detail == null)
            {
                Result.Crocking_Detail = new List<FabricCrkShrkTestCrocking_Detail>();
            }
            Result.MDivisionID = this.MDivisionID;

            Result.Crocking_Main.CrockingTestBeforePicture = Result.Crocking_Main.CrockingTestBeforePicture == null ? null : ImageHelper.ImageCompress(Result.Crocking_Main.CrockingTestBeforePicture);
            Result.Crocking_Main.CrockingTestAfterPicture = Result.Crocking_Main.CrockingTestAfterPicture == null ? null : ImageHelper.ImageCompress(Result.Crocking_Main.CrockingTestAfterPicture);

            BaseResult saveResult = _FabricCrkShrkTest_Service.SaveFabricCrkShrkTestCrockingDetail(Result, this.UserID);
            if (saveResult.Result)
            {
                return RedirectToAction("CrockingTest", new { ID = Result.ID });
            }
            Result.Result = saveResult.Result;
            Result.ErrorMessage = saveResult.ErrorMessage;
            TempData["ModelCrockingTest"] = Result;
            ViewBag.UserMail = this.UserMail;
            return RedirectToAction("CrockingTest", new { ID = Result.ID });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode_Crocking(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.EncodeFabricCrkShrkTestCrockingDetail(ID, this.UserID, out string testResult);
            return Json(new { result.Result, result.ErrorMessage, testResult });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Amend_Crocking(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.AmendFabricCrkShrkTestCrockingDetail(ID);
            return Json(new { result.Result, ErrorMessage = (string.IsNullOrEmpty(result.ErrorMessage) ? string.Empty : result.ErrorMessage.Replace("'", string.Empty)) });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult FailMail_Crocking(long ID, string TO, string CC, string OrderID)
        {
            SendMail_Result result = _FabricCrkShrkTest_Service.SendCrockingFailResultMail(TO, CC, ID, false, OrderID);
            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Report_Crocking(long ID, bool IsToPDF)
        {
            BaseResult result;
            string FileName;
            result = _FabricCrkShrkTest_Service.Crocking_ToExcel(ID, IsToPDF, out FileName);
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult CrockingSendMail(long ID)
        {
            this.CheckSession();

            BaseResult result = null;
            string FileName = string.Empty;

            result = _FabricCrkShrkTest_Service.Crocking_ToExcel(ID, true, out FileName);

            if (!result.Result)
            {
                result.ErrorMessage = result.ErrorMessage.ToString();
            }
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;

            return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, reportPath = reportPath, FileName = FileName });
        }
        #endregion

        #region Heat
        [SessionAuthorizeAttribute]
        public ActionResult HeatTest(long ID)
        {
            FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result = _FabricCrkShrkTest_Service.GetFabricCrkShrkTestHeat_Result(ID);
            if (!fabricCrkShrkTestHeat_Result.Result)
            {
                fabricCrkShrkTestHeat_Result = new FabricCrkShrkTestHeat_Result()
                {
                    Heat_Main = new FabricCrkShrkTestHeat_Main(),
                    Heat_Detail = new List<FabricCrkShrkTestHeat_Detail>(),
                    ErrorMessage = $@"msg.WithInfo('{ (string.IsNullOrEmpty(fabricCrkShrkTestHeat_Result.ErrorMessage) ? string.Empty : fabricCrkShrkTestHeat_Result.ErrorMessage.Replace("'", string.Empty)) }');",
                };
            }

            if (TempData["ModelHeatTest"] != null)
            {
                FabricCrkShrkTestHeat_Result saveResult = (FabricCrkShrkTestHeat_Result)TempData["ModelHeatTest"];
                fabricCrkShrkTestHeat_Result.Heat_Main.HeatRemark = saveResult.Heat_Main.HeatRemark;
                fabricCrkShrkTestHeat_Result.Heat_Detail = saveResult.Heat_Detail;
                fabricCrkShrkTestHeat_Result.Result = saveResult.Result;
                fabricCrkShrkTestHeat_Result.ErrorMessage = $@"msg.WithInfo('{(string.IsNullOrEmpty(saveResult.ErrorMessage) ? string.Empty : saveResult.ErrorMessage.Replace("'", string.Empty))  }');EditMode = true;";
            }

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View(fabricCrkShrkTestHeat_Result);
        }

        [SessionAuthorizeAttribute]
        public ActionResult AddHeatDetailRow(int lastNO, string BrandID)
        {
            string html = string.Empty;
            string defaultOriginalHorizontal = "0";
            string defaultOriginalVertical = "0";

            if (BrandID == "LLL")
            {
                defaultOriginalHorizontal = "30";
                defaultOriginalVertical = "30";
            }

            html += $"<tr idx='{lastNO}' class='row-content' style='vertical-align: middle; text-align: center;'>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputRollSelectItem width6vw' id='Heat_Detail_{lastNO}__Roll' name='Heat_Detail[{lastNO}].Roll' type='text' value='' >";
            html += "<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...' >";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputRollSelectItem width8vw' id='Heat_Detail_{lastNO}__Dyelot' name='Heat_Detail[{lastNO}].Dyelot' type='text' value='' >";
            html += "<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...' >";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalOriginal 必須是數字。' data-val-required='HorizontalOriginal 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalOriginal' name='Heat_Detail[{lastNO}].HorizontalOriginal' step='0.01' type='number' value='{defaultOriginalHorizontal}' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalOriginal 必須是數字。' data-val-required='VerticalOriginal 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalOriginal' name='Heat_Detail[{lastNO}].VerticalOriginal' step='0.01' type='number' value='{defaultOriginalVertical}' onchange='value=QtyCheck(value)'>";
            html += "</td>";

            html += "<td>";
            html += $"<select id='Heat_Detail_{lastNO}__Result' name='Heat_Detail[{lastNO}].Result' class='blue width6vw' onchange='changeResultColor(this)'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest1 必須是數字。' data-val-required='HorizontalTest1 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalTest1' name='Heat_Detail[{lastNO}].HorizontalTest1' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest2 必須是數字。' data-val-required='HorizontalTest2 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalTest2' name='Heat_Detail[{lastNO}].HorizontalTest2' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest3 必須是數字。' data-val-required='HorizontalTest3 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalTest3' name='Heat_Detail[{lastNO}].HorizontalTest3' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalAverage 必須是數字。' data-val-required='HorizontalAverage 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalAverage' name='Heat_Detail[{lastNO}].HorizontalAverage' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalRate 必須是數字。' data-val-required='HorizontalRate 欄位是必要項。' id='Heat_Detail_{lastNO}__HorizontalRate' name='Heat_Detail[{lastNO}].HorizontalRate' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest1 必須是數字。' data-val-required='VerticalTest1 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalTest1' name='Heat_Detail[{lastNO}].VerticalTest1' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest2 必須是數字。' data-val-required='VerticalTest2 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalTest2' name='Heat_Detail[{lastNO}].VerticalTest2' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest3 必須是數字。' data-val-required='VerticalTest3 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalTest3' name='Heat_Detail[{lastNO}].VerticalTest3' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalAverage 必須是數字。' data-val-required='VerticalAverage 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalAverage' name='Heat_Detail[{lastNO}].VerticalAverage' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalRate 必須是數字。' data-val-required='VerticalRate 欄位是必要項。' id='Heat_Detail_{lastNO}__VerticalRate' name='Heat_Detail[{lastNO}].VerticalRate' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='date-picker width9vw' data-val='true' data-val-date='欄位 Inspdate 必須是日期。' id='Heat_Detail_{lastNO}__Inspdate' name='Heat_Detail[{lastNO}].Inspdate' type='text' value='' >";
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
            html += $"<input id='Heat_Detail_{lastNO}__Remark' name='Heat_Detail[{lastNO}].Remark' type='text' value='' >";
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
        [SessionAuthorizeAttribute]
        public ActionResult HeatTestSave(FabricCrkShrkTestHeat_Result Result)
        {
            if (Result.Heat_Detail == null)
            {
                Result.Heat_Detail = new List<FabricCrkShrkTestHeat_Detail>();
            }
            Result.MDivisionID = this.MDivisionID;

            Result.Heat_Main.HeatTestBeforePicture = Result.Heat_Main.HeatTestBeforePicture == null ? null : ImageHelper.ImageCompress(Result.Heat_Main.HeatTestBeforePicture);
            Result.Heat_Main.HeatTestAfterPicture = Result.Heat_Main.HeatTestAfterPicture == null ? null : ImageHelper.ImageCompress(Result.Heat_Main.HeatTestAfterPicture);

            BaseResult saveResult = _FabricCrkShrkTest_Service.SaveFabricCrkShrkTestHeatDetail(Result, this.UserID);

            if (saveResult.Result)
            {
                return RedirectToAction("HeatTest", new { ID = Result.ID });
            }
            Result.Result = saveResult.Result;
            Result.ErrorMessage = saveResult.ErrorMessage;
            TempData["ModelHeatTest"] = Result;
            ViewBag.UserMail = this.UserMail;
            return RedirectToAction("HeatTest", new { ID = Result.ID });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode_Heat(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.EncodeFabricCrkShrkTestHeatDetail(ID, this.UserID, out string testResult);
            return Json(new { result.Result, result.ErrorMessage, testResult });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Amend_Heat(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.AmendFabricCrkShrkTestHeatDetail(ID);
            return Json(new { result.Result, ErrorMessage = (result.ErrorMessage == null ? string.Empty : result.ErrorMessage.Replace("'", string.Empty)) });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult FailMail_Heat(long ID, string TO, string CC, string OrderID)
        {
            SendMail_Result result = _FabricCrkShrkTest_Service.SendHeatFailResultMail(TO, CC, ID, false, OrderID);
            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Report_Heat(long ID)
        {
            BaseResult result;
            result = _FabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestHeatDetail(ID, out string FileName);
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult HeatSendMail(long ID)
        {
            this.CheckSession();

            BaseResult result = null;
            string FileName = string.Empty;

            result = _FabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestHeatDetail(ID, out FileName);

            if (!result.Result)
            {
                result.ErrorMessage = result.ErrorMessage.ToString();
            }
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;

            return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, reportPath = reportPath, FileName = FileName });
        }
        #endregion

        #region Iron
        [SessionAuthorizeAttribute]
        public ActionResult IronTest(long ID)
        {
            FabricCrkShrkTestIron_Result fabricCrkShrkTestIron_Result = _FabricCrkShrkTest_Service.GetFabricCrkShrkTestIron_Result(ID);
            if (!fabricCrkShrkTestIron_Result.Result)
            {
                fabricCrkShrkTestIron_Result = new FabricCrkShrkTestIron_Result()
                {
                    Iron_Main = new FabricCrkShrkTestIron_Main(),
                    Iron_Detail = new List<FabricCrkShrkTestIron_Detail>(),
                    ErrorMessage = $@"msg.WithInfo('{(string.IsNullOrEmpty(fabricCrkShrkTestIron_Result.ErrorMessage) ? string.Empty : fabricCrkShrkTestIron_Result.ErrorMessage.Replace("'", string.Empty))}');",
                };
            }

            if (TempData["ModelIronTest"] != null)
            {
                FabricCrkShrkTestIron_Result saveResult = (FabricCrkShrkTestIron_Result)TempData["ModelIronTest"];
                fabricCrkShrkTestIron_Result.Iron_Main.IronRemark = saveResult.Iron_Main.IronRemark;
                fabricCrkShrkTestIron_Result.Iron_Detail = saveResult.Iron_Detail;
                fabricCrkShrkTestIron_Result.Result = saveResult.Result;
                fabricCrkShrkTestIron_Result.ErrorMessage = $@"msg.WithInfo('{(string.IsNullOrEmpty(saveResult.ErrorMessage) ? string.Empty : saveResult.ErrorMessage.Replace("'", string.Empty))}');EditMode = true;";
            }

            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View(fabricCrkShrkTestIron_Result);
        }

        [SessionAuthorizeAttribute]
        public ActionResult AddIronDetailRow(int lastNO, string BrandID)
        {
            string defaultOriginalHorizontal = "0";
            string defaultOriginalVertical = "0";

            if (BrandID == "LLL")
            {
                defaultOriginalHorizontal = "30";
                defaultOriginalVertical = "30";
            }
            string html = string.Empty;
            html += $"<tr idx='{lastNO}' class='row-content' style='vertical-align: middle; text-align: center;'>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputRollSelectItem width6vw' id='Iron_Detail_{lastNO}__Roll' name='Iron_Detail[{lastNO}].Roll' type='text' value='' >";
            html += "<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...' >";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputRollSelectItem width8vw' id='Iron_Detail_{lastNO}__Dyelot' name='Iron_Detail[{lastNO}].Dyelot' type='text' value='' >";
            html += "<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...' >";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalOriginal 必須是數字。' data-val-required='HorizontalOriginal 欄位是必要項。' id='Iron_Detail_{lastNO}__HorizontalOriginal' name='Iron_Detail[{lastNO}].HorizontalOriginal' step='0.01' type='number' value='{defaultOriginalHorizontal}' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalOriginal 必須是數字。' data-val-required='VerticalOriginal 欄位是必要項。' id='Iron_Detail_{lastNO}__VerticalOriginal' name='Iron_Detail[{lastNO}].VerticalOriginal' step='0.01' type='number' value='{defaultOriginalVertical}' onchange='value=QtyCheck(value)'>";
            html += "</td>";

            html += "<td>";
            html += $"<select id='Iron_Detail_{lastNO}__Result' name='Iron_Detail[{lastNO}].Result' class='blue width6vw' onchange='changeResultColor(this)'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest1 必須是數字。' data-val-required='HorizontalTest1 欄位是必要項。' id='Iron_Detail_{lastNO}__HorizontalTest1' name='Iron_Detail[{lastNO}].HorizontalTest1' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest2 必須是數字。' data-val-required='HorizontalTest2 欄位是必要項。' id='Iron_Detail_{lastNO}__HorizontalTest2' name='Iron_Detail[{lastNO}].HorizontalTest2' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest3 必須是數字。' data-val-required='HorizontalTest3 欄位是必要項。' id='Iron_Detail_{lastNO}__HorizontalTest3' name='Iron_Detail[{lastNO}].HorizontalTest3' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalAverage 必須是數字。' data-val-required='HorizontalAverage 欄位是必要項。' id='Iron_Detail_{lastNO}__HorizontalAverage' name='Iron_Detail[{lastNO}].HorizontalAverage' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalRate 必須是數字。' data-val-required='HorizontalRate 欄位是必要項。' id='Iron_Detail_{lastNO}__HorizontalRate' name='Iron_Detail[{lastNO}].HorizontalRate' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest1 必須是數字。' data-val-required='VerticalTest1 欄位是必要項。' id='Iron_Detail_{lastNO}__VerticalTest1' name='Iron_Detail[{lastNO}].VerticalTest1' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest2 必須是數字。' data-val-required='VerticalTest2 欄位是必要項。' id='Iron_Detail_{lastNO}__VerticalTest2' name='Iron_Detail[{lastNO}].VerticalTest2' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest3 必須是數字。' data-val-required='VerticalTest3 欄位是必要項。' id='Iron_Detail_{lastNO}__VerticalTest3' name='Iron_Detail[{lastNO}].VerticalTest3' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalAverage 必須是數字。' data-val-required='VerticalAverage 欄位是必要項。' id='Iron_Detail_{lastNO}__VerticalAverage' name='Iron_Detail[{lastNO}].VerticalAverage' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalRate 必須是數字。' data-val-required='VerticalRate 欄位是必要項。' id='Iron_Detail_{lastNO}__VerticalRate' name='Iron_Detail[{lastNO}].VerticalRate' readonly='readonly' step='0.01' type='number' value='0' >";
            html += "</td>";
            html += "<td>";
            html += $"<input class='date-picker width9vw' data-val='true' data-val-date='欄位 Inspdate 必須是日期。' id='Iron_Detail_{lastNO}__Inspdate' name='Iron_Detail[{lastNO}].Inspdate' type='text' value='' >";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputInspectorSelectItem' id='Iron_Detail_{lastNO}__Inspector' name='Iron_Detail[{lastNO}].Inspector' type='text' value='' >";
            html += "<input id='btnDetailInspectorSelectItem' type='button' class='site-btn btn-blue btnInspectorSelectItem' style='margin:0;border:0;' value='...' >";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Iron_Detail_{lastNO}__Name' name='Iron_Detail[{lastNO}].Name' readonly='readonly' type='text' value='' >";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Iron_Detail_{lastNO}__Remark' name='Iron_Detail[{lastNO}].Remark' type='text' value='' >";
            html += "</td>";
            html += "<td>";
            html += $"<input id='Iron_Detail_{lastNO}__LastUpdate' name='Iron_Detail[{lastNO}].LastUpdate' readonly='readonly' type='text' value='' >";
            html += "</td>";
            html += "<td>";
            html += "<img class='detailDelete' src='/Image/Icon/Delete.png' style='min-width:30px'>";
            html += "</td>";
            html += "</tr>";
            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult IronTestSave(FabricCrkShrkTestIron_Result Result)
        {
            if (Result.Iron_Detail == null)
            {
                Result.Iron_Detail = new List<FabricCrkShrkTestIron_Detail>();
            }
            Result.MDivisionID = this.MDivisionID;

            Result.Iron_Main.IronTestBeforePicture = Result.Iron_Main.IronTestBeforePicture == null ? null : ImageHelper.ImageCompress(Result.Iron_Main.IronTestBeforePicture);
            Result.Iron_Main.IronTestAfterPicture = Result.Iron_Main.IronTestAfterPicture == null ? null : ImageHelper.ImageCompress(Result.Iron_Main.IronTestAfterPicture);

            BaseResult saveResult = _FabricCrkShrkTest_Service.SaveFabricCrkShrkTestIronDetail(Result, this.UserID);

            if (saveResult.Result)
            {
                return RedirectToAction("IronTest", new { ID = Result.ID });
            }
            Result.Result = saveResult.Result;
            Result.ErrorMessage = saveResult.ErrorMessage;
            TempData["ModelIronTest"] = Result;
            ViewBag.UserMail = this.UserMail;
            return RedirectToAction("IronTest", new { ID = Result.ID });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Encode_Iron(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.EncodeFabricCrkShrkTestIronDetail(ID, this.UserID, out string testResult);
            return Json(new { result.Result, result.ErrorMessage, testResult });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Amend_Iron(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.AmendFabricCrkShrkTestIronDetail(ID);
            return Json(new { result.Result, ErrorMessage = (result.ErrorMessage == null ? string.Empty : result.ErrorMessage.Replace("'", string.Empty)) });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult FailMail_Iron(long ID, string TO, string CC, string OrderID)
        {
            SendMail_Result result = _FabricCrkShrkTest_Service.SendIronFailResultMail(TO, CC, ID, false, OrderID);
            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Report_Iron(long ID)
        {
            BaseResult result;
            result = _FabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestIronDetail(ID, out string FileName);
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult IronSendMail(long ID)
        {
            this.CheckSession();

            BaseResult result = null;
            string FileName = string.Empty;

            result = _FabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestIronDetail(ID, out FileName);

            if (!result.Result)
            {
                result.ErrorMessage = result.ErrorMessage.ToString();
            }
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;

            return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, reportPath = reportPath, FileName = FileName });
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
                    ErrorMessage = $@"msg.WithInfo('{ (string.IsNullOrEmpty(fabricCrkShrkTestWash_Result.ErrorMessage) ? string.Empty : fabricCrkShrkTestWash_Result.ErrorMessage.Replace("'", string.Empty))  }');",
                };
            }

            if (TempData["ModelWashTest"] != null)
            {
                FabricCrkShrkTestWash_Result saveResult = (FabricCrkShrkTestWash_Result)TempData["ModelWashTest"];
                fabricCrkShrkTestWash_Result.Wash_Main.SkewnessOptionID = saveResult.Wash_Main.SkewnessOptionID;
                fabricCrkShrkTestWash_Result.Wash_Main.WashRemark = saveResult.Wash_Main.WashRemark;
                fabricCrkShrkTestWash_Result.Wash_Detail = saveResult.Wash_Detail;
                fabricCrkShrkTestWash_Result.Result = saveResult.Result;
                fabricCrkShrkTestWash_Result.ErrorMessage = $@"msg.WithInfo('{ (string.IsNullOrEmpty(saveResult.ErrorMessage) ? string.Empty : saveResult.ErrorMessage.Replace("'", string.Empty))  }');EditMode = true;";
            }

            List<SelectListItem> skewnessOptionList = new SetListItem().ItemListBinding(skewnessOption);
            ViewBag.ResultList = new SetListItem().ItemListBinding(resultType);
            ViewBag.SkewnessOptionList = skewnessOptionList;
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View(fabricCrkShrkTestWash_Result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult AddWashDetailRow(int lastNO, string BrandID)
        {
            string defaultOriginalHorizontal = "0";
            string defaultOriginalVertical = "0";

            if (BrandID == "LLL")
            {
                defaultOriginalHorizontal = "30";
                defaultOriginalVertical = "30";
            }
            string html = string.Empty;
            html += $"<tr idx='{lastNO}' class='row-content' style='vertical-align: middle; text-align: center;'>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputRollSelectItem width6vw' id='Wash_Detail_{lastNO}__Roll' name='Wash_Detail[{lastNO}].Roll' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input class='inputRollSelectItem width8vw' id='Wash_Detail_{lastNO}__Dyelot' name='Wash_Detail[{lastNO}].Dyelot' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnRollSelectItem' style='margin:0;border:0;' value='...'>";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalOriginal 必須是數字。' data-val-required='HorizontalOriginal 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalOriginal' name='Wash_Detail[{lastNO}].HorizontalOriginal' step='0.01' type='number' value='{defaultOriginalHorizontal}' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalOriginal 必須是數字。' data-val-required='VerticalOriginal 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalOriginal' name='Wash_Detail[{lastNO}].VerticalOriginal' step='0.01' type='number' value='{defaultOriginalVertical}' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<select id='Wash_Main_{lastNO}__Result' name='Wash_Detail[{lastNO}].Result' class='blue width6vw' onchange='changeResultColor(this)'>";
            html += "<option value='Pass'>Pass</option>";
            html += "<option value='Fail'>Fail</option>";
            html += "</select>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest1 必須是數字。' data-val-required='HorizontalTest1 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalTest1' name='Wash_Detail[{lastNO}].HorizontalTest1' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest2 必須是數字。' data-val-required='HorizontalTest2 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalTest2' name='Wash_Detail[{lastNO}].HorizontalTest2' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='HorizontalTest' data-val='true' data-val-number='欄位 HorizontalTest3 必須是數字。' data-val-required='HorizontalTest3 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalTest3' name='Wash_Detail[{lastNO}].HorizontalTest3' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalAverage 必須是數字。' data-val-required='HorizontalAverage 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalAverage' name='Wash_Detail[{lastNO}].HorizontalAverage' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 HorizontalRate 必須是數字。' data-val-required='HorizontalRate 欄位是必要項。' id='Wash_Detail_{lastNO}__HorizontalRate' name='Wash_Detail[{lastNO}].HorizontalRate' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest1 必須是數字。' data-val-required='VerticalTest1 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalTest1' name='Wash_Detail[{lastNO}].VerticalTest1' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest2 必須是數字。' data-val-required='VerticalTest2 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalTest2' name='Wash_Detail[{lastNO}].VerticalTest2' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='VerticalTest' data-val='true' data-val-number='欄位 VerticalTest3 必須是數字。' data-val-required='VerticalTest3 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalTest3' name='Wash_Detail[{lastNO}].VerticalTest3' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalAverage 必須是數字。' data-val-required='VerticalAverage 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalAverage' name='Wash_Detail[{lastNO}].VerticalAverage' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 VerticalRate 必須是數字。' data-val-required='VerticalRate 欄位是必要項。' id='Wash_Detail_{lastNO}__VerticalRate' name='Wash_Detail[{lastNO}].VerticalRate' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest1 必須是數字。' data-val-required='SkewnessTest1 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest1' name='Wash_Detail[{lastNO}].SkewnessTest1' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest2 必須是數字。' data-val-required='SkewnessTest2 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest2' name='Wash_Detail[{lastNO}].SkewnessTest2' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td style=''>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest3 必須是數字。' data-val-required='SkewnessTest3 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest3' name='Wash_Detail[{lastNO}].SkewnessTest3' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td style=''>";
            html += $"<input class='SkewnessTest' data-val='true' data-val-number='欄位 SkewnessTest4 必須是數字。' data-val-required='SkewnessTest4 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessTest4' name='Wash_Detail[{lastNO}].SkewnessTest4' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)'>";
            html += "</td>";
            html += "<td>";
            html += $"<input data-val='true' data-val-number='欄位 SkewnessRate 必須是數字。' data-val-required='SkewnessRate 欄位是必要項。' id='Wash_Detail_{lastNO}__SkewnessRate' name='Wash_Detail[{lastNO}].SkewnessRate' readonly='readonly' step='0.01' type='number' value='0'>";
            html += "</td>";

            html += "<td>";
            html += $"<input class='date-picker width9vw' data-val='true' data-val-date='欄位 Inspdate 必須是日期。' id='Wash_Detail_{lastNO}__Inspdate' name='Wash_Detail[{lastNO}].Inspdate' type='text' value=''>";
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
            html += $"<input id='Wash_Detail_{lastNO}__Remark' name='Wash_Detail[{lastNO}].Remark' type='text' value=''>";
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
        [SessionAuthorizeAttribute]
        public ActionResult WashTestSave(FabricCrkShrkTestWash_Result Result)
        {
            if (Result.Wash_Detail == null)
            {
                Result.Wash_Detail = new List<FabricCrkShrkTestWash_Detail>();
            }
            Result.MDivisionID = this.MDivisionID;

            Result.Wash_Main.WashTestBeforePicture = Result.Wash_Main.WashTestBeforePicture == null ? null : ImageHelper.ImageCompress(Result.Wash_Main.WashTestBeforePicture);
            Result.Wash_Main.WashTestAfterPicture = Result.Wash_Main.WashTestAfterPicture == null ? null : ImageHelper.ImageCompress(Result.Wash_Main.WashTestAfterPicture);

            BaseResult saveResult = _FabricCrkShrkTest_Service.SaveFabricCrkShrkTestWashDetail(Result, this.UserID);
            if (saveResult.Result)
            {
                return RedirectToAction("WashTest", new { ID = Result.ID });
            }

            Result.Result = saveResult.Result;
            Result.ErrorMessage = saveResult.ErrorMessage;
            TempData["ModelWashTest"] = Result;
            ViewBag.UserMail = this.UserMail;
            return RedirectToAction("WashTest", new { ID = Result.ID });
        }

        [HttpPost]
        public JsonResult Encode_Wash(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.EncodeFabricCrkShrkTestWashDetail(ID, this.UserID, out string testResult);
            return Json(new { result.Result, result.ErrorMessage, testResult });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Amend_Wash(long ID)
        {
            BaseResult result = _FabricCrkShrkTest_Service.AmendFabricCrkShrkTestWashDetail(ID);
            return Json(new { result.Result, ErrorMessage = (result.ErrorMessage == null ? string.Empty : result.ErrorMessage.Replace("'", string.Empty)) });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult FailMail_Wash(long ID, string TO, string CC, string OrderID)
        {
            SendMail_Result result = _FabricCrkShrkTest_Service.SendWashFailResultMail(TO, CC, ID, false, OrderID);
            return Json(result);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Report_Wash(long ID)
        {
            BaseResult result;
            result = _FabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestWashDetail(ID, out string FileName);
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult WashSendMail(long ID)
        {
            this.CheckSession();

            BaseResult result = null;
            string FileName = string.Empty;

            result = _FabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestWashDetail(ID, out FileName);

            if (!result.Result)
            {
                result.ErrorMessage = result.ErrorMessage.ToString();
            }
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;

            return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, reportPath = reportPath, FileName = FileName });
        }
        #endregion
    }
}