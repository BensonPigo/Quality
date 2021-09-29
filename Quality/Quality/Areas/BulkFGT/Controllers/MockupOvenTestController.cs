using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
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
    public class MockupOvenTestController : BaseController
    {
        private IMockupOvenService _MockupOvenService;

        public MockupOvenTestController()
        {
            _MockupOvenService = new MockupOvenService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.MockupOvenTest,,";
        }

        // GET: BulkFGT/MockupOvenTest
        public ActionResult Index()
        {
            MockupOven_ViewModel model = new MockupOven_ViewModel()
            {
                MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                ReportNo_Source = new List<string>(),
                Request = new MockupOven_Request(),
            };

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source;
            ViewBag.ArtworkTypeID_Source = new SetListItem().ItemListBinding(new List<string>());
            ViewBag.AccessoryRefNo_Source = new SetListItem().ItemListBinding(new List<string>());
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            MockupOven_ViewModel Req = new MockupOven_ViewModel()
            {
                Request = new MockupOven_Request()
                {
                    ReportNo = ReportNo,
                    BrandID = BrandID,
                    SeasonID = SeasonID,
                    StyleID = StyleID,
                    Article = Article,
                }
            };

            MockupOven_ViewModel model = _MockupOvenService.GetMockupOven(Req.Request);


            if (model == null)
            {
                model = new MockupOven_ViewModel()
                {
                    MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                    ReportNo_Source = new List<string>(),
                    ErrorMessage = $"msg.WithInfo('No Data Found');",
                };
            }


            model.Request = Req.Request;
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(MockupOven_ViewModel Req)
        {
            MockupOven_ViewModel model = _MockupOvenService.GetMockupOven(Req.Request);

            if (model == null)
            {
                model = new MockupOven_ViewModel()
                {
                    MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                    ReportNo_Source = new List<string>(),
                    ErrorMessage = $"msg.WithInfo('No Data Found');",
                };
            }


            model.Request = Req.Request;
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", model);
        }


        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult NewSave(MockupOven_ViewModel Req)
        {
            if (Req.MockupOven_Detail == null)
            {
                Req.MockupOven_Detail = new List<MockupOven_Detail_ViewModel>();
            }

            BaseResult result = _MockupOvenService.Create(Req, this.MDivisionID, this.UserID, out string ReportNo);

            Req.Request = new MockupOven_Request()
            {
                BrandID = Req.BrandID,
                SeasonID = Req.SeasonID,
                StyleID = Req.StyleID,
                Article = Req.Article,
                ReportNo = ReportNo,
            };

            MockupOven_ViewModel model = _MockupOvenService.GetMockupOven(Req.Request);
            if (model == null)
            {
                model = new MockupOven_ViewModel()
                {
                    MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                    ReportNo_Source = new List<string>(),
                    Request = new MockupOven_Request(),
                };
            }

            if (!result.Result)
            {
                model.ErrorMessage = $"msg.WithInfo('" + result.ErrorMessage.ToString().Replace("\r\n", "<br />") + "');EditMode=true;";
            }
            else if (result.Result && model.Result == "Fail")
            {
                model.ErrorMessage = "FailMail();";
            }

            model.Request = Req.Request;
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", model);
        }


        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult EditSave(MockupOven_ViewModel Req)
        {
            if (Req.MockupOven_Detail == null)
            {
                Req.MockupOven_Detail = new List<MockupOven_Detail_ViewModel>();
            }

            BaseResult result = _MockupOvenService.Update(Req, this.UserID);
            Req.Request = new MockupOven_Request()
            {
                BrandID = Req.BrandID,
                SeasonID = Req.SeasonID,
                StyleID = Req.StyleID,
                Article = Req.Article,
                ReportNo = Req.ReportNo,
            };

            MockupOven_ViewModel model = _MockupOvenService.GetMockupOven(Req.Request);

            if (model == null)
            {
                model = new MockupOven_ViewModel()
                {
                    MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                    ReportNo_Source = new List<string>(),
                    Request = new MockupOven_Request(),
                };
            }

            Req.Result = model.Result;
            Req.MRName = model.MRName;
            Req.MRMail = model.MRMail;
            Req.TechnicianName = model.TechnicianName;
            Req.LastEditName = model.LastEditName;
            Req.MockupOven_Detail = model.MockupOven_Detail;
            if (!result.Result)
            {
                Req.ErrorMessage = $"msg.WithInfo('" + result.ErrorMessage.ToString().Replace("\r\n", "<br />") + "');EditMode=true;";
            }
            else if (result.Result && model.Result == "Fail")
            {
                Req.ErrorMessage = "FailMail();";
            }

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", Req);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult DeleteReportNo(MockupOven_ViewModel Req)
        {
            Req.ReportNo = Req.Request.ReportNo;
            BaseResult result = _MockupOvenService.Delete(Req);
            Req.Request.ReportNo = "";
            MockupOven_ViewModel model = _MockupOvenService.GetMockupOven(Req.Request);

            if (model == null)
            {
                model = new MockupOven_ViewModel()
                {
                    MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                    ReportNo_Source = new List<string>(),
                    Request = new MockupOven_Request(),
                };
            }
            if (!result.Result)
            {
                model.ErrorMessage = $"msg.WithInfo('" + result.ErrorMessage.ToString().Replace("\r\n", "<br />") + "');";
            }
            model.Request = Req.Request;

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", model);
        }

        [HttpPost]
        public ActionResult ToPDF(MockupOven_Request mockupOven_Request)
        {
            this.CheckSession();
            MockupOven_ViewModel model = _MockupOvenService.GetMockupOven(mockupOven_Request);
            if (model == null)
            {
                return Json(new { Result = false, ErrorMessage = "msg.WithInfo('No Data Found');" });
            }

            Report_Result report_Result = _MockupOvenService.GetPDF(model);
            string tempFilePath = report_Result.TempFileName;
            tempFilePath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + tempFilePath;
            if (!report_Result.Result)
            {
                report_Result.ErrorMessage = report_Result.ErrorMessage.ToString().Replace("\r\n", "<br />");
            }
            
            return Json(new { Result = report_Result.Result, ErrorMessage = report_Result.ErrorMessage, reportPath = tempFilePath, FileName = report_Result.TempFileName });
        }

        [HttpPost]
        public JsonResult SPBlur(string POID)
        {
            string BrandID = string.Empty;
            string SeasonID = string.Empty;
            string StyleID = string.Empty;
            string Article = string.Empty;

            Orders order = new Orders();
            order.ID = POID;
            List<Orders> orderResult = _MockupOvenService.GetOrders(order);

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
                List<Order_Qty> order_qtyResult = _MockupOvenService.GetDistinctArticle(order_qty);

                if (order_qtyResult.Count == 1)
                {
                    Article = order_qtyResult.FirstOrDefault().Article;
                }

            }

            return Json(new { ErrMsg = "", BrandID = BrandID, SeasonID = SeasonID, StyleID = StyleID, Article = Article });
        }

        [HttpPost]
        public ActionResult GetArtworkTypeID_Source(string BrandID, string SeasonID, string StyleID)
        {
            return Json(GetArtworkTypeIDList(BrandID, SeasonID, StyleID));
        }

        private List<SelectListItem> GetArtworkTypeIDList(string BrandID, string SeasonID, string StyleID)
        {
            StyleArtwork_Request styleArtwork_Request = new StyleArtwork_Request()
            {
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
            };
            return _MockupOvenService.GetArtworkTypeID(styleArtwork_Request);
        }

        public ActionResult GetAccessoryRefNo_Source(string BrandID, string SeasonID, string StyleID)
        {
            return Json(GetAccessoryRefNoList(BrandID, SeasonID, StyleID));
        }

        private List<SelectListItem> GetAccessoryRefNoList(string BrandID, string SeasonID, string StyleID)
        {
            AccessoryRefNo_Request AccessoryRefNo_Request = new AccessoryRefNo_Request()
            {
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
            };
            return _MockupOvenService.GetAccessoryRefNo(AccessoryRefNo_Request);
        }

        [HttpPost]
        public ActionResult AddDetailRow(int lastNO,string BrandID, string SeasonID, string StyleID)
        {
            List<SelectListItem> AccessoryRefNo_Source= new SetListItem().ItemListBinding(new List<string>()); 

            if (!string.IsNullOrWhiteSpace(BrandID) && !string.IsNullOrWhiteSpace(SeasonID) && !string.IsNullOrWhiteSpace(StyleID))
            {
              AccessoryRefNo_Source = GetAccessoryRefNoList(BrandID, SeasonID, StyleID);
            }

 
            MockupOven_ViewModel model = new MockupOven_ViewModel();

            string html = "";
            html += $"<tr idx='{lastNO}'>";
            html += $"<td><input id='Seq{lastNO}' idx='{lastNO}' type ='hidden'></input> <input id='MockupOven_Detail_{lastNO}__TypeofPrint' name='MockupOven_Detail[{lastNO}].TypeofPrint' class='OnlyEdit' type='text' value=''></td>";
            html += $"<td><input id='MockupOven_Detail_{lastNO}__Design' name='MockupOven_Detail[{lastNO}].Design' class='OnlyEdit' type='text' ></td>"; 
            html += $"<td><div class='input-group'><input id='MockupOven_Detail_{lastNO}__ArtworkColor' name='MockupOven_Detail[{lastNO}].ArtworkColor'  class ='AFColor' type='hidden'><input id='MockupOven_Detail_{lastNO}__ArtworkColorName' name='MockupOven_Detail[{lastNO}].ArtworkColorName'  class ='AFColor' type='text' readonly='readonly'> <input idv='{lastNO}' type='button' class='btnArtworkColorItem  site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>";
            html += $"<td><select id='MockupOven_Detail_{lastNO}__AccessoryRefNo_Source' name='MockupOven_Detail[{lastNO}].AccessoryRefNo_Source'  class='OnlyEdit' style='width: 157px;'><option value=''></option>"; 
            foreach (var val in AccessoryRefNo_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";
            html += $"<td><input id='MockupOven_Detail_{lastNO}__FabricRefNo' name='MockupOven_Detail[{lastNO}].FabricRefNo' type='text' ></td>";
            html += $"<td><div class='input-group'><input id='MockupOven_Detail_{lastNO}__FabricColor' name='MockupOven_Detail[{lastNO}].FabricColor'  class ='AFColor' type='hidden'><input id='MockupOven_Detail_{lastNO}__FabricColorName' name='MockupOven_Detail[{ lastNO}].FabricColorName'  class ='AFColor' type='text' readonly='readonly'> <input idv='{lastNO}' type='button' class='btnFabricColorItem  site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>";

            html += $"<td><select id='MockupOven_Detail_{lastNO}__Result' name='MockupOven_Detail[{lastNO}].Result' class='OnlyEdit result blue' onchange='changeResult()' ><option value=''></option>"; 
            foreach (var val in model.Result_Source)
            {
                if (val.Value == "Pass")
                {
                    html += "<option value='" + val.Value + "' SELECTED>" + val.Text + "</option>";
                }
                else
                {
                    html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
                }

            }
            html += "</select></td>";

            html += $"<td><input id='MockupOven_Detail_{lastNO}__SCIRefno' name='MockupOven_Detail[{lastNO}].Remark' type='text' class='OnlyEdit'></td>"; 
            html += $"<td><input id='MockupOven_Detail_{lastNO}__ColorID' name='MockupOven_Detail[{lastNO}].LastUpdate' type='text'  readonly='readonly' ></td>";

            html += "<td> <div style='width: 5vw;'><img  class='detailDelete' src='/Image/Icon/Delete.png' width='30'> </div></td>";
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

            SendMail_Result result = _MockupOvenService.FailSendMail(mail);
            return Json(result);
        }
    }
}