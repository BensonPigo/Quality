using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using DatabaseObject;
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
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.MockupCrocking,,";
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
            ViewBag.FactoryID = this.FactoryID;
            return View(Result);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(MockupCrocking_ViewModel Req)
        {
            //example:ADIDAS, 19FW, F1915KYB012, ED5770
            
            MockupCrocking_ViewModel mockupCrocking_ViewModel = _MockupCrockingService.GetMockupCrocking(Req.Request);

            //若沒錯誤訊息 且 MockupCrocking_ViewModel == null, 則是沒資料
            if (mockupCrocking_ViewModel == null)
            {
                mockupCrocking_ViewModel = new MockupCrocking_ViewModel()
                {
                    ErrorMessage = $"msg.WithInfo('No Data Found');",
                    MockupCrocking_Detail = new List<MockupCrocking_Detail_ViewModel>(),
                    ReportNo_Source = new List<string>(),
                };
            }

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(mockupCrocking_ViewModel.ReportNo_Source);
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            return View("Index", mockupCrocking_ViewModel);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult New(MockupCrocking_ViewModel Req)
        {
            Req.AddName = this.UserID;
            BaseResult result = _MockupCrockingService.Create(Req, this.MDivisionID);
            MockupCrocking_ViewModel mockupCrocking_ViewModel = _MockupCrockingService.GetMockupCrocking(Req.Request);
            if (mockupCrocking_ViewModel == null)
            {
                mockupCrocking_ViewModel = new MockupCrocking_ViewModel() {  ReportNo_Source = new List<string>(), };
            }

            Req.ReportNo = mockupCrocking_ViewModel.ReportNo;
            Req.LastEditName = mockupCrocking_ViewModel.LastEditName;
            if (!result.Result)
            {
                Req.ErrorMessage = $"msg.WithInfo('" + result.ErrorMessage.ToString().Replace("\r\n", "<br />") + "');";
            }
            else if (result.Result && mockupCrocking_ViewModel.Result == "Fail")
            {
                Req.ErrorMessage = "FailMail();";
            }

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(mockupCrocking_ViewModel.ReportNo_Source);
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            return View("Index", Req);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult Edit(MockupCrocking_ViewModel Req)
        {
            Req.EditName = this.UserID;
            foreach (var item in Req.MockupCrocking_Detail)
            {
                item.EditName = this.UserID;
            }
            BaseResult result = _MockupCrockingService.Update(Req);
            MockupCrocking_ViewModel mockupCrocking_ViewModel = _MockupCrockingService.GetMockupCrocking(Req.Request);

            if (mockupCrocking_ViewModel == null)
            {
                mockupCrocking_ViewModel = new MockupCrocking_ViewModel() { ReportNo_Source = new List<string>(), };
            }

            Req.ReportNo = mockupCrocking_ViewModel.ReportNo;
            Req.LastEditName = mockupCrocking_ViewModel.LastEditName;
            if (!result.Result)
            {
                Req.ErrorMessage = $"msg.WithInfo('" + result.ErrorMessage.ToString().Replace("\r\n", "<br />") + "');";
            }
            else if (result.Result && mockupCrocking_ViewModel.Result == "Fail")
            {
                Req.ErrorMessage = "FailMail();";
            }

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(mockupCrocking_ViewModel.ReportNo_Source);
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            return View("Index", Req);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(MockupCrocking_ViewModel Req)
        {
            BaseResult result = _MockupCrockingService.Delete(Req);
            MockupCrocking_ViewModel mockupCrocking_ViewModel = _MockupCrockingService.GetMockupCrocking(Req.Request);

            if (mockupCrocking_ViewModel == null)
            {
                mockupCrocking_ViewModel = new MockupCrocking_ViewModel() { ReportNo_Source = new List<string>(), };
            }

            Req.ReportNo = mockupCrocking_ViewModel.ReportNo;
            Req.LastEditName = mockupCrocking_ViewModel.LastEditName;
            if (!result.Result)
            {
                Req.ErrorMessage = $"msg.WithInfo('" + result.ErrorMessage.ToString().Replace("\r\n", "<br />") + "');";
            }

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(mockupCrocking_ViewModel.ReportNo_Source);
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            return View("Index", Req);
        }

        [HttpPost]
        public JsonResult ToPDF(MockupCrocking_Request mockupCrocking_Request)
        {
            this.CheckSession();
            MockupCrocking_ViewModel mockupCrocking_ViewModel = _MockupCrockingService.GetMockupCrocking(mockupCrocking_Request);
            BaseResult result = new BaseResult();
            if (mockupCrocking_ViewModel == null)
            {
                return Json(new { Result = false, ErrorMessage = "msg.WithInfo('No Data Found');" });
            }

            Report_Result report_Result = _MockupCrockingService.GetPDF(mockupCrocking_ViewModel);
            string tempFilePath = report_Result.TempFileName;
            tempFilePath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + tempFilePath;

            return Json(new { Result = true, reportPath = tempFilePath, FileName = report_Result.TempFileName });
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

        [HttpPost]
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

            html += $"<select id='MockupCrocking_Detail_{lastNo}__DryScale' name='MockupCrocking_Detail[{lastNo}].DryScale' class='NotEdit' style='width:157px;'>";
            html += string.Join("", scaleOption);
            html += "</select >";

            html += "</td>";
            html += "<td>";

            html += $"<select id='MockupCrocking_Detail_{lastNo}__WetScale' name='MockupCrocking_Detail[{lastNo}].WetScale' class='NotEdit' style='width:157px;'>";
            html += string.Join("", scaleOption);
            html += "</select>";

            html += "<td>";

            html += $"<select id='MockupCrocking_Detail_{lastNo}__Result' class='result' name='MockupCrocking_Detail[{lastNo}].Result' style='width:157px;' onchange='changeResult()'>";
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

        [HttpPost]
        public JsonResult FailMail(string ReportNo, string TO, string CC)
        {
            MockupFailMail_Request mail = new MockupFailMail_Request()
            {
                ReportNo = ReportNo,
                To = TO,
                CC = CC,
            };

            SendMail_Result result = _MockupCrockingService.FailSendMail(mail);
            return Json(result);
        }
    }
}