using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using Quality.Helper;
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
        [SessionAuthorizeAttribute]
        public ActionResult Index()
        {
            MockupOven_ViewModel model = new MockupOven_ViewModel()
            {
                MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                ReportNo_Source = new List<string>(),
                Request = new MockupOven_Request(),
                ScaleID_Source = _MockupOvenService.GetScale(),
            };

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source;
            ViewBag.ChangeScaleList = model.ScaleID_Source;
            ViewBag.ResultChangeList = model.Result_Source;
            ViewBag.StainingScaleList = model.ScaleID_Source;
            ViewBag.ResultStainList = model.Result_Source;
            ViewBag.ArtworkTypeID_Source = new SetListItem().ItemListBinding(new List<string>());
            ViewBag.AccessoryRefNo_Source = new SetListItem().ItemListBinding(new List<string>());
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        [SessionAuthorizeAttribute]
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
                    ErrorMessage = $@"msg.WithInfo(""No Data Found"");",
                    ScaleID_Source = _MockupOvenService.GetScale(),
                };
            }

            if (!model.ReturnResult)
            {
                model = new MockupOven_ViewModel()
                {
                    MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                    ReportNo_Source = new List<string>(),
                    ErrorMessage = $@"msg.WithInfo(""{ (string.IsNullOrEmpty(model.ErrorMessage) ? string.Empty : model.ErrorMessage.Replace("'", string.Empty).Replace("\r\n", "<br />")) }"");",
                    ScaleID_Source = _MockupOvenService.GetScale(),
                };
            }

            model.Request = Req.Request;
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source;
            ViewBag.ChangeScaleList = model.ScaleID_Source;
            ViewBag.ResultChangeList = model.Result_Source;
            ViewBag.StainingScaleList = model.ScaleID_Source;
            ViewBag.ResultStainList = model.Result_Source;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
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
                    ErrorMessage = $@"msg.WithInfo(""No Data Found"");",
                    ScaleID_Source = _MockupOvenService.GetScale(),
                };
            }

            if (!model.ReturnResult)
            {
                model = new MockupOven_ViewModel()
                {
                    MockupOven_Detail = new List<MockupOven_Detail_ViewModel>(),
                    ReportNo_Source = new List<string>(),
                    ErrorMessage = $@"msg.WithInfo(""{ (string.IsNullOrEmpty(model.ErrorMessage) ? string.Empty : model.ErrorMessage.Replace("\r\n", "<br />")) }"");",
                    ScaleID_Source = _MockupOvenService.GetScale(),
                };
            }

            model.Request = Req.Request;
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source;
            ViewBag.ChangeScaleList = model.ScaleID_Source;
            ViewBag.ResultChangeList = model.Result_Source;
            ViewBag.StainingScaleList = model.ScaleID_Source; 
            ViewBag.ResultStainList = model.Result_Source;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewData["AccessoryRefNo_Source"] = GetAccessoryRefNoList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult NewSave(MockupOven_ViewModel Req)
        {
            if (Req.MockupOven_Detail == null)
            {
                Req.MockupOven_Detail = new List<MockupOven_Detail_ViewModel>();
            }

            Req.TestBeforePicture = Req.TestBeforePicture == null ? null : ImageHelper.ImageCompress(Req.TestBeforePicture);
            Req.TestAfterPicture = Req.TestAfterPicture == null ? null : ImageHelper.ImageCompress(Req.TestAfterPicture);

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
                    ScaleID_Source = _MockupOvenService.GetScale(),
                };
            }   

            if (!result.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo('{ (string.IsNullOrEmpty(result.ErrorMessage) ? string.Empty : result.ErrorMessage.Replace("\r\n", "<br />")) }');EditMode=true;";
            }
            else if (result.Result && model.Result == "Fail")
            {
                model.ErrorMessage = "FailMail();";
            }

            model.Request = Req.Request;
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source;
            ViewBag.ChangeScaleList = model.ScaleID_Source;
            ViewBag.ResultChangeList = model.Result_Source;
            ViewBag.StainingScaleList = model.ScaleID_Source;
            ViewBag.ResultStainList = model.Result_Source;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", model);
        }


        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult EditSave(MockupOven_ViewModel Req)
        {
            if (Req.MockupOven_Detail == null)
            {
                Req.MockupOven_Detail = new List<MockupOven_Detail_ViewModel>();
            }

            Req.TestBeforePicture = Req.TestBeforePicture == null ? null : ImageHelper.ImageCompress(Req.TestBeforePicture);
            Req.TestAfterPicture = Req.TestAfterPicture == null ? null : ImageHelper.ImageCompress(Req.TestAfterPicture);

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
                    ScaleID_Source = _MockupOvenService.GetScale(),
                };
            }

            Req.MailSubject = model.MailSubject;
            Req.Result = model.Result;
            Req.MRName = model.MRName;
            Req.MRMail = model.MRMail;
            Req.TechnicianName = model.TechnicianName;
            Req.LastEditName = model.LastEditName;
            Req.MockupOven_Detail = model.MockupOven_Detail;
            if (!result.Result)
            {
                Req.ErrorMessage = $@"msg.WithInfo('{result.ErrorMessage.Replace("'", string.Empty).Replace("\r\n", "<br />")}');EditMode=true;";
            }
            else if (result.Result && model.Result == "Fail")
            {
                Req.ErrorMessage = "FailMail();";
            }

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source;
            ViewBag.ChangeScaleList = model.ScaleID_Source;
            ViewBag.ResultChangeList = model.Result_Source;
            ViewBag.StainingScaleList = model.ScaleID_Source;
            ViewBag.ResultStainList = model.Result_Source;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", Req);
        }

        [SessionAuthorizeAttribute]
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
                    ScaleID_Source = _MockupOvenService.GetScale(),
                };
            }
            if (!result.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo('{result.ErrorMessage.Replace("'", string.Empty).Replace("\r\n", "<br />")}');";
            }
            model.Request = Req.Request;

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source;
            ViewBag.ChangeScaleList = model.ScaleID_Source;
            ViewBag.ResultChangeList = model.Result_Source;
            ViewBag.StainingScaleList = model.ScaleID_Source;
            ViewBag.ResultStainList = model.Result_Source;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(model.Request.BrandID, model.Request.SeasonID, model.Request.StyleID);
            ViewBag.FactoryID = this.FactoryID;
            ViewBag.UserMail = this.UserMail;
            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        public ActionResult ToPDF(MockupOven_Request mockupOven_Request)
        {
            this.CheckSession();
            MockupOven_ViewModel model = _MockupOvenService.GetMockupOven(mockupOven_Request);
            if (model == null)
            {
                return Json(new { Result = false, ErrorMessage = @"msg.WithInfo(""No Data Found"");" });
            }

            Report_Result report_Result = _MockupOvenService.GetPDF(model);
            string tempFilePath = report_Result.TempFileName;
            tempFilePath = "/TMP/" + tempFilePath;
            if (!report_Result.Result)
            {
                report_Result.ErrorMessage = report_Result.ErrorMessage.ToString();
            }
            
            return Json(new { Result = report_Result.Result, ErrorMessage = report_Result.ErrorMessage, reportPath = tempFilePath, FileName = report_Result.TempFileName });
        }

        [SessionAuthorizeAttribute]
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

        [SessionAuthorizeAttribute]
        [HttpPost]
        public JsonResult GetArtwrok_Ajax(string POID, string BrandID, string SeasonID, string StyleID)
        {
            List<SelectListItem> datas = new List<SelectListItem>();
            
            if (!string.IsNullOrEmpty(POID))
            {
                datas = this.GetArtworkTypeIDListByPOID(POID);
            }
            else
            {
                datas = this.GetArtworkTypeIDList(BrandID, SeasonID, StyleID);
            }

            return Json(datas);
        }

        [SessionAuthorizeAttribute]
        private List<SelectListItem> GetArtworkTypeIDList(string BrandID, string SeasonID, string StyleID)
        {
            if (string.IsNullOrEmpty(BrandID) || string.IsNullOrEmpty(SeasonID) || string.IsNullOrEmpty(StyleID))
            {
                return new List<SelectListItem>();
            }

            StyleArtwork_Request styleArtwork_Request = new StyleArtwork_Request()
            {
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
            };
            return _MockupOvenService.GetArtworkTypeID(styleArtwork_Request);
        }

        [SessionAuthorizeAttribute]
        private List<SelectListItem> GetArtworkTypeIDListByPOID(string POID)
        {
            StyleArtwork_Request styleArtwork_Request = new StyleArtwork_Request()
            {
                POID = POID,
            };
            return _MockupOvenService.GetArtworkTypeID(styleArtwork_Request);
        }


        [SessionAuthorizeAttribute]
        public ActionResult GetAccessoryRefNo_Source(string BrandID, string SeasonID, string StyleID)
        {
            return Json(GetAccessoryRefNoList(BrandID, SeasonID, StyleID));
        }

        [SessionAuthorizeAttribute]
        private List<SelectListItem> GetAccessoryRefNoList(string BrandID, string SeasonID, string StyleID)
        {
            if (string.IsNullOrEmpty(BrandID) || string.IsNullOrEmpty(SeasonID) || string.IsNullOrEmpty(StyleID))
            {
                return new List<SelectListItem>();
            }

            AccessoryRefNo_Request AccessoryRefNo_Request = new AccessoryRefNo_Request()
            {
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
            };
            return _MockupOvenService.GetAccessoryRefNo(AccessoryRefNo_Request);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        public ActionResult AddDetailRow(int lastNO,string BrandID, string SeasonID, string StyleID)
        {
            List<SelectListItem> AccessoryRefNo_Source= new SetListItem().ItemListBinding(new List<string>());
            List<SelectListItem> Scales = _MockupOvenService.GetScale();

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
            html += $"<td><select id='MockupOven_Detail_{lastNO}__AccessoryRefNo' name='MockupOven_Detail[{lastNO}].AccessoryRefNo'  class='OnlyEdit AccessoryRefno' style='width: 157px;'>"; 
            foreach (var val in AccessoryRefNo_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";
            html += $"<td><input id='MockupOven_Detail_{lastNO}__FabricRefNo' name='MockupOven_Detail[{lastNO}].FabricRefNo' type='text' ></td>";
            html += $"<td><div class='input-group'><input id='MockupOven_Detail_{lastNO}__FabricColor' name='MockupOven_Detail[{lastNO}].FabricColor'  class ='AFColor' type='hidden'><input id='MockupOven_Detail_{lastNO}__FabricColorName' name='MockupOven_Detail[{ lastNO}].FabricColorName'  class ='AFColor' type='text' readonly='readonly'> <input idv='{lastNO}' type='button' class='btnFabricColorItem  site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>";
            html += $"<td><input  readonly='readonly'  id='MockupOven_Detail_{lastNO}__Result' name='MockupOven_Detail[{lastNO}].Result' class='blue' type='text' value='Pass'></td>"; // Result

            html += $"<td><select id='MockupOven_Detail_{lastNO}__ChangeScale' name='MockupOven_Detail[{lastNO}].ChangeScale'>"; // ChangeScale
            foreach (var val in Scales)
            {
                string selected = string.Empty;
                if (val.Value == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            html += $"<td><select onchange='selectChange(this)' id='MockupOven_Detail_{lastNO}__ResultChange' name='MockupOven_Detail[{lastNO}].ResultChange' class='blue'>"; // ResultChange
            foreach (var val in model.Result_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            html += $"<td><select id='MockupOven_Detail_{lastNO}__StainingScale' name='MockupOven_Detail[{lastNO}].StainingScale'>"; // StainingScale
            foreach (var val in Scales)
            {
                string selected = string.Empty;
                if (val.Value == "4-5")
                {
                    selected = "selected";
                }
                html += $"<option {selected} value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            html += $"<td><select onchange='selectChange(this)' id='MockupOven_Detail_{lastNO}__ResultStain' name='MockupOven_Detail[{lastNO}].ResultStain' class='blue'>"; // ResultStain
            foreach (var val in model.Result_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            html += $"<td><input id='MockupOven_Detail_{lastNO}__SCIRefno' name='MockupOven_Detail[{lastNO}].Remark' type='text' class='OnlyEdit'></td>"; 
            html += $"<td><input id='MockupOven_Detail_{lastNO}__ColorID' name='MockupOven_Detail[{lastNO}].LastUpdate' type='text'  readonly='readonly' ></td>";

            html += "<td> <div style='width: 5vw;'><img  class='detailDelete' src='/Image/Icon/Delete.png' width='30'> </div></td>";
            html += "</tr>";

            return Content(html);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        public JsonResult SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            MockupFailMail_Request mail = new MockupFailMail_Request()
            {
                ReportNo = ReportNo,
                To = TO,
                CC = CC,
                Subject = Subject,
                Body = Body,
                Files = Files,
            };

            SendMail_Result result = _MockupOvenService.SendMail(mail);
            return Json(result);
        }
    }
}