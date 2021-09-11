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

            BusinessLogicLayer.Interface.BulkFGT.IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
            MockupCrocking_ViewModel mockupCrocking_ViewModel = _MockupCrockingService.GetMockupCrocking(Req.Request);

            //若沒錯誤訊息 且 MockupCrocking_ViewModel == null, 則是沒資料

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(mockupCrocking_ViewModel.ReportNo_Source);
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            return View("Index", mockupCrocking_ViewModel);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult NewSave(MockupCrocking_ViewModel Req)
        {
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

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult DeleteReportNo(MockupCrocking_ViewModel Req)
        {
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(Req.ReportNo_Source);
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);

            return View("Index", Req);
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

        [HttpPost]
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

        [HttpPost]
        public JsonResult SPBlur(string POID)
        {
            string BrandID = string.Empty;
            string SeasonID = string.Empty;
            string StyleID = string.Empty;
            string Article = string.Empty;

            IMockupCrockingService mockupCrockingService = new MockupCrockingService();

            Orders order = new Orders();
            order.ID = POID;
            List<Orders> orderResult = mockupCrockingService.GetOrders(order);

            if (orderResult.Count == 0)
            {
                return Json(new { ErrMsg = $"Cannot found SP# {POID}." });
            }

            if (orderResult.Count == 1)
            {
                BrandID = orderResult.FirstOrDefault().BrandID;
                SeasonID = orderResult.FirstOrDefault().SeasonID;
                StyleID = orderResult.FirstOrDefault().StyleID;

                Order_Qty order_qty = new Order_Qty();
                order_qty.ID = POID;
                List<Order_Qty> order_qtyResult = mockupCrockingService.GetDistinctArticle(order_qty);

                if (order_qtyResult.Count == 1)
                {
                    Article = order_qtyResult.FirstOrDefault().Article;
                }
            }

            return Json(new { ErrMsg = "", BrandID = BrandID, SeasonID = SeasonID, StyleID = StyleID, Article = Article });
        }

        public ActionResult AddDetailRow(int lastNo)
        {
            MockupCrocking_ViewModel mockupCrocking_ViewModel = new MockupCrocking_ViewModel();
            List<string> resultOption = new List<string>();
            foreach (var item in mockupCrocking_ViewModel.Result_Source)
            {
                resultOption.Add($"<option value='{item.Value}'>{item.Text}</option>");
            }

            List<string> scaleOption = new List<string>();
            foreach (var item in mockupCrocking_ViewModel.Scale_Source)
            {
                scaleOption.Add($"<option value='{item.Value}'>{item.Text}</option>");
            }

            lastNo = lastNo + 1;
            string html = string.Empty;
            html += $"<tr idx='{lastNo}' class='row-content' style='vertical-align: middle; text-align: center;'>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input id='MockupCrocking_Detail_{lastNo}__Design' name='MockupCrocking_Detail[{lastNo}].Design' type='text' value=''>";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input id='MockupCrocking_Detail_{lastNo}__ArtworkColor' class='AFColor' name='MockupCrocking_Detail[{lastNo}].ArtworkColor' type='hidden' value=''>";
            html += $"<input id='MockupCrocking_Detail_{lastNo}__ArtworkColorName' class='AFColor' name='MockupCrocking_Detail[{lastNo}].ArtworkColorName' readonly='readonly' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnArtworkColorItem' style='margin:0;border:0;' value='...'>";
            html += "</div>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='MockupCrocking_Detail_{lastNo}__FabricRefNo' name='MockupCrocking_Detail[{lastNo}].FabricRefNo' type='text' value=''>";
            html += "</td>";
            html += "<td>";
            html += "<div class='input-group'>";
            html += $"<input id='MockupCrocking_Detail_{lastNo}__FabricColor' class='AFColor' name='MockupCrocking_Detail[{lastNo}].FabricColor' type='hidden' value=''>";
            html += $"<input id='MockupCrocking_Detail_{lastNo}__FabricColorName' class='AFColor' name='MockupCrocking_Detail[{lastNo}].FabricColorName' readonly='readonly' type='text' value=''>";
            html += $"<input type='button' class='site-btn btn-blue btnFabricColorItem' style='margin:0;border:0;' value='...'>";
            html += "</div>";
            html += "</td>";
            html += "<td>";

            html += $"<select id='MockupCrocking_Detail_{lastNo}__DryScale' name='Model.MockupCrocking_Detail[{lastNo}].DryScale' class='NotEdit' style='width:157px;'>";
            html += string.Join("", scaleOption);
            html += "</select >";

            html += "</td>";
            html += "<td>";

            html += $"<select id='MockupCrocking_Detail_{lastNo}__WetScale' name='Model.MockupCrocking_Detail[{lastNo}].WetScale' class='NotEdit' style='width:157px;'>";
            html += string.Join("", scaleOption);
            html += "</select>";

            html += "<td>";

            html += $"<select id='MockupCrocking_Detail_{lastNo}__Result' class='result' name='Model.MockupCrocking_Detail[{lastNo}].Result' style='width:157px;' onchange='changeResult()'>";
            html += string.Join("", resultOption);
            html += "</select>";

            html += "</td >";
            html += "<td>";
            html += $"<input id='MockupCrocking_Detail_{lastNo}__Remark' name='MockupCrocking_Detail[{lastNo}].Remark' type='text' value=''>";
            html += "</td>";
            html += "<td>";
            html += $"<input id='MockupCrocking_Detail_{lastNo}__LastUpdate' name='MockupCrocking_Detail[{lastNo}].LastUpdate' readonly='readonly' type='text' value=''>";
            html += "</td>";
            html += "<td>";
            html += "<img class='detailDelete' src='/Image/Icon/Delete.png' width='30' style='min-width:30px'>";
            html += "</td>";
            html += "</tr>";
            return Content(html);
        }

 
    }
}