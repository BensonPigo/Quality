using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
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
        private List<string> TestResultmm = new List<string>() { "<=4", ">4" };
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
                garmentTest_Details = new List<GarmentTest_Detail_ViewModel>() { 
                    new GarmentTest_Detail_ViewModel() { ID = 1, No = 1 },
                    new GarmentTest_Detail_ViewModel() { ID = 1, No = 2 },
                } ,
                req = new GarmentTest_Request(), 
            };

            List<SelectListItem> SizeCodeList = new SetListItem().ItemListBinding(Result.SizeCodes);
            List<SelectListItem> MtlTypeIDList = new SetListItem().ItemListBinding(this.MtlTypeIDs);

            ViewBag.SizeCodeList = SizeCodeList;
            ViewBag.MtlTypeIDList = MtlTypeIDList;
            return View(Result);
        }

        [HttpPost]
        public ActionResult Index(GarmentTest_Request Req)
        {
            Req.Factory = this.FactoryID;
            GarmentTest_Result Result = _GarmentTest_Service.GetGarmentTest(Req);
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
            return View(Result);
        }

        [HttpPost]
        public JsonResult SaveDetail(List<GarmentTest_Detail> details)
        {
            var result = new { Result = false, ErrMsg = "Err" };
            return Json(result);
        }

        [HttpPost]
        public ActionResult AddDetailRow(string ID, int lastNO, string OrderID, string Article)
        {
            int i = lastNO - 1;
            GarmentTest_Detail_ViewModel detail = new GarmentTest_Detail_ViewModel();
            List<string> sizecodes = _GarmentTest_Service.Get_SizeCode(OrderID, Article);
            string html = "";
            html += "<tr>";
            html += "<td><a href='/BulkFGT/GarmentTest/Detail/" + ID + "?No=" + lastNO.ToString() + "' idx='" + ID + "' idv = '" + lastNO.ToString() + "'>" + lastNO.ToString() + "</a></td>";
            html += "<td><input id='garmentTest_Details_" + i + "_OrderID' name='garmentTest_Details[" + i + "].OrderID' type='text'></td>";
            html += "<td><select id='garmentTest_Details_" + i + "_SizeCode' name='garmentTest_Details[" + i + "].SizeCode'><option value=''></option>";
            foreach(string val in sizecodes)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";
            html += "<td><input class='form-control date-picker hasDatepicker' data-val='true' data-val-date='欄位 檢驗日期 必須是日期。' id='garmentTest_Details_" + i + "_inspdate' name='garmentTest_Details[" + i + "].inspdate' type='text' value=''></td>";
            html += "<td><select id='garmentTest_Details_" + i + "_MtlTypeID' name='garmentTest_Details[" + i + "].MtlTypeID'><option value=''></option>";
            foreach (string val in MtlTypeIDs)
            {
                html += "<option value='" + val + "'>" + val + "</option>";
            }
            html += "</select></td>";
            html += "<td class='red'>Fail</td>";
            html += "<td><input id='garmentTest_Details_" + i + "_NonSeamBreakageTest' name='garmentTest_Details[" + i + "].NonSeamBreakageTest' type='checkbox'><input name='garmentTest_Details.NonSeamBreakageTest' type='hidden'></td>";
            html += "<td></td>";
            html += "<td></td>";
            html += "<td></td>";
            html += "<td><input id='garmentTest_Details_" + i + "_inspector' name='garmentTest_Details[" + i + "].inspector' type='text'></td>";
            html += "<td></td>";
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
            html += "<td><img class='detailDelete display-None' src='/Image/Icon/Delete.png' width='30'></td>";
            html += "</tr>";

            return Content(html);
        }

        public ActionResult Detail(string ID, string No)
        {
            List<SelectListItem> TemperatureList = new SetListItem().ItemListBinding(Temperatures);
            List<SelectListItem> MachineList = new SetListItem().ItemListBinding(Machines);
            List<SelectListItem> NeckList = new SetListItem().ItemListBinding(Necks);
            List<SelectListItem> WashList = new SetListItem().ItemListBinding(Washs);
            List<SelectListItem> TestResultPassList = new SetListItem().ItemListBinding(TestResultPass);
            List<SelectListItem> TestResultmmList = new SetListItem().ItemListBinding(TestResultmm);

            GarmentTest_Detail_Result Detail_Result = new GarmentTest_Detail_Result()
            {
                Scales = new List<string>(),
                Main = new GarmentTest_ViewModel(),
                Detail = new GarmentTest_Detail_ViewModel(),
                Shrinkages = new List<GarmentTest_Detail_Shrinkage>(),
                Spiralities = new List<Garment_Detail_Spirality>(),
                Apperance = new List<GarmentTest_Detail_Apperance_ViewModel>(),
                FGWT = new List<GarmentTest_Detail_FGWT_ViewModel>(),
                FGPT = new List<GarmentTest_Detail_FGPT_ViewModel>(),
            };

            List<SelectListItem> ScaleList = new SetListItem().ItemListBinding(Detail_Result.Scales);

            ViewBag.TemperatureList = TemperatureList;
            ViewBag.MachineList = MachineList;
            ViewBag.NeckList = NeckList;
            ViewBag.WashList = WashList;
            ViewBag.ScaleList = ScaleList;
            ViewBag.TestResultPassList = TestResultPassList;
            ViewBag.TestResultmmList = TestResultmmList;
            return View(Detail_Result);
        }
    }
}