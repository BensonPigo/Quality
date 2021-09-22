using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class GarmentTestController : BaseController
    {
        private IGarmentTest_Service _GarmentTest_Service;
        private List<string> MtlTypeIDs = new List<string>() { "KNIT", "WOVEN" };
        private List<object> Temperatures = new List<object>() { 30, 40, 50, 60 };
        private List<string> Machines = new List<string>() { "Top Load", "Front Load" };
        private Dictionary<string, object> Necks = new Dictionary<string, object>() {{ "Yes", true }, { "No", false }};
        private List<string> Washs = new List<string>() { "N/A", "Accepted", "Rejected" };
        private List<string> TestResultPass = new List<string>() { "Pass", "Fail" };
        private List<string> TestResultmm = new List<string>()  { "<=4", ">4" };
        public GarmentTestController()
        {
            _GarmentTest_Service = new GarmentTest_Service();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.GarmentTest,,";
        }

        // GET: BulkFGT/GarmentTest
        public ActionResult Index()
        {
            GarmentTest_Result Result = new GarmentTest_Result()
            {
                Result = true,
                SizeCodes = _GarmentTest_Service.Get_SizeCode(string.Empty, string.Empty),
                garmentTest = new GarmentTest_ViewModel(),
                garmentTest_Details = new List<GarmentTest_Detail_ViewModel>(),
                req = new GarmentTest_Request(), 
            };

            List<SelectListItem> SizeCodeList = new SetListItem().ItemListBinding(Result.SizeCodes);
            List<SelectListItem> MtlTypeIDList = new SetListItem().ItemListBinding(this.MtlTypeIDs);

            ViewBag.SizeCodeList = SizeCodeList;
            ViewBag.MtlTypeIDList = MtlTypeIDList;
            return View(Result);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexBack(string Brand, string Season, string Style, string Article)
        {
            GarmentTest_Request Req = new GarmentTest_Request()
            {
                Brand = Brand,
                Season = Season,
                Style = Style,
                Article = Article,
                Factory = this.FactoryID,
            }; 
            GarmentTest_Result Result = _GarmentTest_Service.GetGarmentTest(Req);
            if (!Result.Result)
            {
                Result.garmentTest = new GarmentTest_ViewModel() { ID = 0 };
                Result.SizeCodes = new List<string>();
            }

            if (Result.garmentTest_Details == null || Result.garmentTest_Details.Count == 0)
            {
                Result.garmentTest_Details = new List<GarmentTest_Detail_ViewModel>()
                {
                    new GarmentTest_Detail_ViewModel() { No = 1, ID = Result.garmentTest.ID },
                };
            }

            Result.req = Req;
            List<SelectListItem> SizeCodeList = new SetListItem().ItemListBinding(Result.SizeCodes);
            List<SelectListItem> MtlTypeIDList = new SetListItem().ItemListBinding(this.MtlTypeIDs);
            ViewBag.SizeCodeList = SizeCodeList;
            ViewBag.MtlTypeIDList = MtlTypeIDList;
            ViewBag.GarmentTestRequest = Req;
            return View("Index", Result);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>

        [HttpPost]
        public ActionResult Index(GarmentTest_Request Req)
        {
            Req.Factory = this.FactoryID;
            GarmentTest_Result Result = _GarmentTest_Service.GetGarmentTest(Req);
            if (!Result.Result)
            {
                Result.garmentTest = new GarmentTest_ViewModel() { ID = 0 };
                Result.SizeCodes = new List<string>();
            }
            
            //if (Result.garmentTest_Details == null || Result.garmentTest_Details.Count == 0)
            //{
            //    Result.garmentTest_Details = new List<GarmentTest_Detail_ViewModel>()
            //    {
            //        new GarmentTest_Detail_ViewModel() { No = 1, ID = Result.garmentTest.ID },
            //    };
            //}

            Result.req = Req;
            List<SelectListItem> SizeCodeList = new SetListItem().ItemListBinding(Result.SizeCodes);
            List<SelectListItem> MtlTypeIDList = new SetListItem().ItemListBinding(this.MtlTypeIDs);
            ViewBag.SizeCodeList = SizeCodeList;
            ViewBag.MtlTypeIDList = MtlTypeIDList;
            ViewBag.GarmentTestRequest = Req;
            return View(Result);
        }

        [HttpPost]
        public JsonResult SaveDetail(GarmentTest_ViewModel main, List<GarmentTest_Detail> details)
        {            
            GarmentTest_ViewModel result = _GarmentTest_Service.Save_GarmentTest(main, details, this.UserID);

            GarmentTest_Result result1 = new GarmentTest_Result()
            {
                Result = result.SaveResult,
                ErrMsg = result.ErrMsg,
                req = new GarmentTest_Request()
                {
                    Style = main.StyleID,
                    Article = main.Article,
                    Brand = main.BrandID,
                    Season = main.SeasonID,
                }
            };

            return Json(result1);
        }

        [HttpPost]
        public JsonResult DeleteDetail(string ID, string No)
        {
            GarmentTest_ViewModel result = _GarmentTest_Service.DeleteDetail(ID, No);

            return Json(result);
        }

        [HttpPost]
        public JsonResult SendMail(string ID, string No)
        {
            GarmentTest_ViewModel result = _GarmentTest_Service.SendMail(ID, No, this.UserID);
            GarmentTest_Detail_ViewModel detail = _GarmentTest_Service.Get_Detail(ID, No);
            result.Sender = detail.Sender;
            result.SendDate = detail.SendDate.HasValue ? detail.SendDate.Value.ToString("yyyy/MM/dd HH:mm:ss") : string.Empty;

            return Json(result);
        }

        [HttpPost]
        public JsonResult ReceiveMail(string ID, string No)
        {
            GarmentTest_ViewModel result = _GarmentTest_Service.ReceiveMail(ID, No, this.UserID);
            GarmentTest_Detail_ViewModel detail = _GarmentTest_Service.Get_Detail(ID, No);
            result.Sender = detail.Receiver;
            result.SendDate = detail.ReceiveDate.HasValue ? detail.ReceiveDate.Value.ToString("yyyy/MM/dd HH:mm:ss") : string.Empty;

            return Json(result);
        }

        [HttpPost]
        public ActionResult AddDetailRow(string ID, int lastNO, string OrderID, string Article, string Brand, string Season, string Style)
        {
            int i = lastNO - 1;
            GarmentTest_Detail_ViewModel detail = new GarmentTest_Detail_ViewModel();

            bool chk = _GarmentTest_Service.CheckOrderID(OrderID, Brand, Season, Style);
            List<string> sizecodes = new List<string>();
            if (chk) {
                sizecodes = _GarmentTest_Service.Get_SizeCode(OrderID, Article);
            }
            else {
                sizecodes = _GarmentTest_Service.Get_SizeCode(Style, Season, Brand);
            }

            string html = "";
            html += "<tr idx='" + i + "'>";
            html += "<td><a idx='" + ID + "' idv = '" + lastNO.ToString() + "'>" + lastNO.ToString() + "</a></td>";
            html += "<td><input id='garmentTest_Details_" + i + "_OrderID' name='garmentTest_Details[" + i + "].OrderID' class='Detail_OrderID' type='text'></td>";
            html += "<td><select id='garmentTest_Details_" + i + "_SizeCode' name='garmentTest_Details[" + i + "].SizeCode' class='Detail_SizeCode'><option value=''></option>";
            foreach(string val in sizecodes)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";
            html += "<td><input class='form-control date-picker' id='garmentTest_Details_" + i + "_inspdate' name='garmentTest_Details[" + i + "].inspdate' type='text' value=''></td>";
            html += "<td><select id='garmentTest_Details_" + i + "_MtlTypeID' name='garmentTest_Details[" + i + "].MtlTypeID'><option value=''></option>";
            foreach (string val in MtlTypeIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";
            html += "<td></td>";
            html += "<td><input id='garmentTest_Details_" + i + "_NonSeamBreakageTest' name='garmentTest_Details[" + i + "].NonSeamBreakageTest' type='checkbox' class='bigSize'></td>";
            html += "<td></td>";
            html += "<td></td>";
            html += "<td></td>";
            html += "<td><input id='garmentTest_Details_" + i + "_inspector' name='garmentTest_Details[" + i + "].inspector' type='text'><input id='btnDetailInspectorSelectItem' type='button' class='site-btn btn-blue' style='margin:0;border:0;' value='...'></td>";
            html += "<td><input id='garmentTest_Details_" + i + "__GarmentTest_Detail_Inspector' name='garmentTest_Details[" + i + "].GarmentTest_Detail_Inspector' readonly='readonly' style='width: 8vw' type='text'></td>";
            html += "<td><input  id='garmentTest_Details_" + i + "_Remark' name='garmentTest_Details[" + i + "].Remark' type='text'></td>";
            html += "<td><img class='SendMail display-None' src='/Image/Icon/Mail.png' width='30'></td>";
            html += "<td></td>";
            html += "<td></td>";
            html += "<td><img class='ReceiveMail display-None' src='/Image/Icon/Mail.png' width='30'></td>";
            html += "<td></td>";
            html += "<td></td>";
            html += "<td></td>";
            html += "<td></td>";
            html += "<td><img class='detailEdit display-None' src='/Image/Icon/Edit.png' width='30'></td>";
            html += "<td><img class='detailDelete' src='/Image/Icon/Delete.png' width='30'></td>";
            html += "</tr>";

            return Content(html);
        }

        [HttpPost]
        public ActionResult ChangeSizeCode(string OrderID, string Brand, string Season, string Style, string Article)
        {
            bool chk = _GarmentTest_Service.CheckOrderID(OrderID, Brand, Season, Style);
            string html = "";
            List<string> sizeCodes = new List<string>();
            if (chk)
            {
                 sizeCodes = _GarmentTest_Service.Get_SizeCode(OrderID, Article);
            }
            else
            {
                sizeCodes = _GarmentTest_Service.Get_SizeCode(Style, Season, Brand); 
            }
            foreach (string val in sizeCodes)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }

            return Content(html);
        }

        public ActionResult Detail(string ID, string No, bool EditMode)
        {
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> MachineList = new SetListItem().ItemListBinding(Machines);
            List<SelectListItem> NeckList = new SetListItem().ItemListBinding(Necks);
            List<SelectListItem> WashList = new SetListItem().ItemListBinding(Washs);
            List<SelectListItem> TestResultPassList = new SetListItem().ItemListBinding(TestResultPass);
            List<SelectListItem> TestResultmmList = new SetListItem().ItemListBinding(TestResultmm);

            GarmentTest_Detail_Result Detail_Result = _GarmentTest_Service.Get_All_Detail(ID, No);
            Detail_Result.EditMode = EditMode;

            List<SelectListItem> ScaleList = new SetListItem().ItemListBinding(Detail_Result.Scales);

            ViewBag.TemperatureList = TemperatureList;
            ViewBag.MachineList = MachineList;
            ViewBag.NeckList = NeckList;
            ViewBag.WashList = WashList;
            ViewBag.ScaleList = ScaleList;
            ViewBag.TestResultPassList = TestResultPassList;
            ViewBag.TestResultmmList = TestResultmmList;
            ViewBag.FactoryID = this.FactoryID;
            return View(Detail_Result);
        }

        [HttpPost]
        public ActionResult DetailSave(GarmentTest_Detail_Result result)
        {
            result.Detail.LineDry = false;
            result.Detail.TumbleDry = false;
            result.Detail.HandWash = false;
            switch (result.Detail.DrySelect)
            {
                case "LineDry":
                    result.Detail.LineDry = true;
                    break;
                case "TumbleDry":
                    result.Detail.TumbleDry = true;
                    break;
                case "HandWash":
                    result.Detail.HandWash = true;
                    break;
            }

            result.Detail.Above50NaturalFibres = false;
            result.Detail.Above50SyntheticFibres = false;
            switch (result.Detail.Above50)
            {
                case "Above50NaturalFibres":
                    result.Detail.Above50NaturalFibres = true;
                    break;
                case "Above50SyntheticFibres":
                    result.Detail.Above50SyntheticFibres = true;
                    break;
            }


            foreach(var item in result.FGPT.Where(x => string.IsNullOrEmpty(x.Location)))
            {
                item.Location = string.Empty;
            }

            GarmentTest_Detail_Result saveresult = _GarmentTest_Service.Save_GarmentTestDetail(result);
            if (saveresult.Result.Value)
            {
                return RedirectToAction("Detail", new { ID = result.Detail.ID.ToString(), No = result.Detail.No.ToString(), EditMode = false });
            }

            GarmentTest_Detail_Result Detail_Result = _GarmentTest_Service.Get_All_Detail(result.Detail.ID.ToString(), result.Detail.No.ToString());
            Detail_Result.EditMode = result.EditMode;
            Detail_Result.Result = saveresult.Result;
            Detail_Result.ErrMsg = saveresult.ErrMsg;

            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> MachineList = new SetListItem().ItemListBinding(Machines);
            List<SelectListItem> NeckList = new SetListItem().ItemListBinding(Necks);
            List<SelectListItem> WashList = new SetListItem().ItemListBinding(Washs);
            List<SelectListItem> TestResultPassList = new SetListItem().ItemListBinding(TestResultPass);
            List<SelectListItem> TestResultmmList = new SetListItem().ItemListBinding(TestResultmm);
            List<SelectListItem> ScaleList = new SetListItem().ItemListBinding(Detail_Result.Scales);
            ViewBag.TemperatureList = TemperatureList;
            ViewBag.MachineList = MachineList;
            ViewBag.NeckList = NeckList;
            ViewBag.WashList = WashList;
            ViewBag.ScaleList = ScaleList;
            ViewBag.TestResultPassList = TestResultPassList;
            ViewBag.TestResultmmList = TestResultmmList;
            ViewBag.FactoryID = this.FactoryID;

            return View("Detail", Detail_Result);
        }

        [HttpPost]
        public JsonResult GenerateFGWT(string ID, string No)
        {
            GarmentTest_ViewModel main = _GarmentTest_Service.Get_Main(ID);
            GarmentTest_Detail_ViewModel detail = _GarmentTest_Service.Get_Detail(ID, No);
            GarmentTest_Result result = _GarmentTest_Service.Generate_FGWT(main, detail);
            return Json(new { result.Result, result.ErrMsg });
        }

        [HttpPost]
        public JsonResult Encode_Detail(string ID, string No, GarmentTest_Service.DetailStatus status)
        {
            GarmentTest_Detail_Result result = _GarmentTest_Service.Encode_Detail(ID, No, status);
            return Json(new { result.Result, result.ErrMsg, result.sentMail });
        }

        [HttpPost]
        public JsonResult FailMail(string ID, string No, string TO, string CC)
        {
            List<Quality_MailGroup> mailGroups = new List<Quality_MailGroup>() {
                new Quality_MailGroup() { ToAddress = TO, CcAddress = CC, }
            };
            GarmentTest_Result result = _GarmentTest_Service.SentMail(ID, No, mailGroups);

            return Json(new { result.Result, result.ErrMsg });
        }

        [HttpPost]
        public JsonResult Report(string ID, string No, GarmentTest_Service.ReportType type, bool IsToPDF)
        {
            GarmentTest_Detail_Result result = _GarmentTest_Service.ToReport(ID, No, type, IsToPDF);
            result.reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + result.reportPath;
            return Json(new { result.Result, result.ErrMsg, result.reportPath });
        }
    }
}