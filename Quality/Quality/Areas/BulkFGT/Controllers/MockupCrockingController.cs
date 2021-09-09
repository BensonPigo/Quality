using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
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
    public class MockupCrockingController : BaseController
    {
        private IMockupCrockingService _MockupCrockingService;
        private MailTools _SendMail;
        public MockupCrockingController()
        {
            _MockupCrockingService = new MockupCrockingService();
            _SendMail = new MailTools();
        }

        // GET: BulkFGT/MockupCrocking
        public ActionResult Index()
        {
            MockupCrocking_ViewModel Result = new MockupCrocking_ViewModel()
            {
                Request = new MockupCrocking_Request
                {
                    BrandID = string.Empty,
                    SeasonID = string.Empty,
                    StyleID = string.Empty,
                    Article = string.Empty,
                },

                MockupCrocking_Detail = new List<MockupCrocking_Detail_ViewModel>(),
                ReportNo_Source = new List<string>(),
            };

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(Result.ReportNo_Source);
            ViewBag.ArtworkTypeID_Source = new SetListItem().ItemListBinding(new List<string>());
            return View(Result);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(MockupCrocking_ViewModel Req)
        {
            //example:ADIDAS, 19FW, F1915KYB012, ED5770
            //因為SMSURUF18BJM07M 不存在，但查得結果，這邊重設修件
            //Req.Request.BrandID = "ADIDAS";
            //Req.Request.SeasonID = "18FW";
            //Req.Request.StyleID = "SMSURUF18BJM07M";
            //Req.Request.Article = "DY3851";

            BusinessLogicLayer.Interface.BulkFGT.IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
            MockupCrocking_ViewModel mockupCrocking_ViewModel = _MockupCrockingService.GetMockupCrocking(Req.Request);

            //若沒錯誤訊息 且 MockupCrocking_ViewModel == null, 則是沒資料

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(mockupCrocking_ViewModel.ReportNo_Source);
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            return View("Index", mockupCrocking_ViewModel);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult EditSave(MockupCrocking_ViewModel Req)
        {
            BusinessLogicLayer.Interface.BulkFGT.IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
            MockupCrocking_ViewModel mockupCrocking_ViewModel = _MockupCrockingService.GetMockupCrocking(Req.Request);

            //若沒錯誤訊息 且 MockupCrocking_ViewModel == null, 則是沒資料

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(mockupCrocking_ViewModel.ReportNo_Source);
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            return View("Index", mockupCrocking_ViewModel);
        }

        /// <summary>
        /// 使用Microsoft.Office.Interop.Excel的寫法
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "ToPDF")]
        public ActionResult ToPDF()
        {
            this.CheckSession();
            try
            {
                MockupCrocking_ViewModel mockupCrocking = (MockupCrocking_ViewModel)TempData["Model"];
                // 1. 在Service層取得資料，生成Excel檔案，放在暫存路徑，回傳檔名
                var result = _MockupCrockingService.GetPDF(mockupCrocking);
                string tempFilePath = result.TempFileName;
                // 2. 取得hotst name，串成下載URL ，傳到準備前端下載
                // URL範例：https://misap:1880/TMP/CFT Comments20210826f7f4ad14-186f-451a-9bc1-6edbcaf6cd65.xlsx 
                // (暫存檔檔名是CFT Comments20210826f7f4ad14-186f-451a-9bc1-6edbcaf6cd65.xlsx)
                tempFilePath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + tempFilePath;

                // 3. 前端下載方式：請參考Index.cshtml的 「window.location.href = '@download'」;

                TempData["tempFilePath"] = tempFilePath;
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult GetArtworkTypeID_Source(string brandID, string seasonID, string styleID)
        {
            return Json(GetArtworkTypeIDList(brandID, seasonID, styleID));
        }

        private List<SelectListItem> GetArtworkTypeIDList(string brandID, string seasonID, string styleID)
        {
            IMockupCrockingService mockupCrockingService = new MockupCrockingService();
            StyleArtwork_Request styleArtwork_Request = new StyleArtwork_Request()
            {
                BrandID = brandID,
                SeasonID = seasonID,
                StyleID = styleID,
            };
            return mockupCrockingService.GetArtworkTypeID(styleArtwork_Request);
        }
    }
}