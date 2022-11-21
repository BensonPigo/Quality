using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
using NPOI.SS.Formula.Functions;
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
    public class HeatTransferWashController : BaseController
    {
        private HeatTransferWashService _Service;
        public HeatTransferWashController()
        {
            _Service = new HeatTransferWashService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.HeatTransferWash,,";
        }

        // GET: BulkFGT/HeatTransferWash
        public ActionResult Index()
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Request = new HeatTransferWash_Request(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            HeatTransferWash_Request Req = new HeatTransferWash_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Request = Req,
                Main = new HeatTransferWash_Result(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            model = _Service.GetHeatTransferWash(Req);


            if (model.Result && !model.ReportNo_Source.Any())
            {
                model.ErrorMessage = "Data not found.";
            }
            else if (!model.Result)
            {
                string ErrorMessage = model.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
                model.ReportNo_Source = new List<SelectListItem>();
            }

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(HeatTransferWash_ViewModel Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Request = Req.Request,
                Main = new HeatTransferWash_Result(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            model = _Service.GetHeatTransferWash(Req.Request);

            if (model.Result && !model.ReportNo_Source.Any())
            {
                model.ErrorMessage = "Data not found.";
            }
            else if (!model.Result)
            {
                string ErrorMessage = model.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
                model.ReportNo_Source = new List<SelectListItem>();
            }

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult NewSave(HeatTransferWash_ViewModel Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Request = new HeatTransferWash_Request(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };


            if (Req.Details == null)
            {
                Req.Details = new List<HeatTransferWash_Detail_Result>();
            }


            Req.Main.TestBeforePicture = Req.Main.TestBeforePicture == null ? null : ImageHelper.ImageCompress(Req.Main.TestBeforePicture);
            Req.Main.TestAfterPicture = Req.Main.TestAfterPicture == null ? null : ImageHelper.ImageCompress(Req.Main.TestAfterPicture);
            Req.Main.AddName = this.UserID;

            BaseResult result = _Service.Create(Req, this.MDivisionID, this.UserID, out string NewReportNo);

            if (!result.Result)
            {
                string ErrorMessage = result.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
            }
            else if (result.Result)
            {
                model = _Service.GetHeatTransferWash(new HeatTransferWash_Request()
                {
                    ReportNo = NewReportNo,
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });
            }

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult EditSave(HeatTransferWash_ViewModel Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Request = new HeatTransferWash_Request(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };


            if (Req.Details == null)
            {
                Req.Details = new List<HeatTransferWash_Detail_Result>();
            }


            Req.Main.TestBeforePicture = Req.Main.TestBeforePicture == null ? null : ImageHelper.ImageCompress(Req.Main.TestBeforePicture);
            Req.Main.TestAfterPicture = Req.Main.TestAfterPicture == null ? null : ImageHelper.ImageCompress(Req.Main.TestAfterPicture);
            Req.Main.AddName = this.UserID;

            BaseResult result = _Service.Update(Req, this.UserID);

            if (!result.Result)
            {
                string ErrorMessage = result.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
            }
            else if (result.Result)
            {
                model = _Service.GetHeatTransferWash(new HeatTransferWash_Request()
                {
                    ReportNo = Req.Main.ReportNo,
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });
            }

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(HeatTransferWash_ViewModel Req)
        {

            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Request = Req.Request,
                Details = new List<HeatTransferWash_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            BaseResult result = _Service.Delete(Req);

            if (!result.Result)
            {
                string ErrorMessage = result.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
            }
            else if (result.Result)
            {
                model = _Service.GetHeatTransferWash(new HeatTransferWash_Request()
                {
                    BrandID = Req.Request.BrandID,
                    SeasonID = Req.Request.SeasonID,
                    StyleID = Req.Request.StyleID,
                    Article = Req.Request.Article,
                });
            }

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Encode")]
        public ActionResult Encode(HeatTransferWash_ViewModel Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel();
            Req.Main = new HeatTransferWash_Result()
            {
                ReportNo = Req.Request.ReportNo,
                Status = "Confirmed",
                EditName = this.UserID,
            };
            BaseResult result = _Service.EncodeAmend(Req);

            if (!result.Result)
            {
                string ErrorMessage = result.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
            }
            else if (result.Result)
            {
                model = _Service.GetHeatTransferWash(new HeatTransferWash_Request()
                {
                    ReportNo = Req.Request.ReportNo,
                    BrandID = Req.Request.BrandID,
                    SeasonID = Req.Request.SeasonID,
                    StyleID = Req.Request.StyleID,
                    Article = Req.Request.Article,
                });
            }

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Amend")]
        public ActionResult Amend(HeatTransferWash_ViewModel Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel();
            Req.Main = new HeatTransferWash_Result()
            {
                ReportNo = Req.Request.ReportNo,
                Status = "New",
                EditName = this.UserID,
            };
            BaseResult result = _Service.EncodeAmend(Req);

            if (!result.Result)
            {
                string ErrorMessage = result.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
            }
            else if (result.Result)
            {
                model = _Service.GetHeatTransferWash(new HeatTransferWash_Request()
                {
                    ReportNo = Req.Request.ReportNo,
                    BrandID = Req.Request.BrandID,
                    SeasonID = Req.Request.SeasonID,
                    StyleID = Req.Request.StyleID,
                    Article = Req.Request.Article,
                });
            }

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        public JsonResult SPBlur(string OrderID)
        {
            string BrandID = string.Empty;
            string SeasonID = string.Empty;
            string StyleID = string.Empty;
            string Article = string.Empty;
            bool Teamwear = false;

            Orders order = new Orders();
            order.ID = OrderID;
            List<Orders> orderResult = _Service.GetOrders(order);

            if (orderResult.Count == 0)
            {
                return Json(new { ErrMsg = $"Cannot found SP# {OrderID}." });
            }

            if (orderResult.Count == 1)
            {
                BrandID = orderResult.FirstOrDefault().BrandID;
                SeasonID = orderResult.FirstOrDefault().SeasonID;
                StyleID = orderResult.FirstOrDefault().StyleID;
                Teamwear = orderResult.FirstOrDefault().Teamwear;

                Order_Qty order_qty = new Order_Qty();
                order_qty.ID = OrderID;
                List<Order_Qty> order_qtyResult = _Service.GetDistinctArticle(order_qty);

                if (order_qtyResult.Any())
                {
                    Article = order_qtyResult.FirstOrDefault().Article;
                }

            }

            return Json(new { ErrMsg = "", BrandID = BrandID, SeasonID = SeasonID, StyleID = StyleID, Article = Article , Teamwear = Teamwear });
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        public ActionResult AddDetailRow(int lastNO, string BrandID, string SeasonID, string StyleID)
        {
            List<SelectListItem> resultSelect = new List<SelectListItem>()
            {
                new SelectListItem(){Text="Pass",Value="Pass"},
                new SelectListItem(){Text="Fail",Value="Fail"},
            };


            MockupOven_ViewModel model = new MockupOven_ViewModel();

            string html = "";
            html += $"<tr idx='{lastNO}'>";

            html += $"<td> <input id='Seq{lastNO}' idx='{lastNO}' type ='hidden'> ";
            html += $"<input id='Details_{lastNO}__FabricRefNo' name='Details[{lastNO}].FabricRefNo' class='OnlyEdit FabricRefNoTxt' type='text' value=''  style='width:85%;'>";
            html += $"<input targetID='Details_{lastNO}__FabricRefNo' type='button' class='btnRefnoSelectItem site-btn btn-blue' style='margin:0;border:0;' value='...' />";
            html += $"</td>";


            html += $"<td> <input id='Details_{lastNO}__HTRefNo' name='Details[{lastNO}].HTRefNo' class='OnlyEdit HTRefNoTxt' type='text' value=''  style='width:85%;'>";
            html += $"     <input targetID='Details_{lastNO}__HTRefNo' type='button' class='btnHTRefNoSelectItem site-btn btn-blue' style='margin:0;border:0;' value='...' />";
            html += $"</td>";
            html += $"<td><select id='Details_{lastNO}__Result' name='Details[{lastNO}].Result'  class='OnlyEdit' >";
            foreach (var val in resultSelect)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";

            html += $"<td><input id='Details_{lastNO}__Remark' name='Details[{lastNO}].Remark' type='text' maxlength='300' class='OnlyEdit'></td>";
            html += $"<td></td>";

            html += "<td> <div style='width: 100%;'><img  class='detailDelete' src='/Image/Icon/Delete.png' width='30'> </div></td>";
            html += "</tr>";

            return Content(html);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult Report(string ReportNo, bool IsToPDF)
        {
            BaseResult result;
            string FileName;
            if (IsToPDF)
            {
                result = _Service.ToReport(ReportNo, IsToPDF, out FileName);
            }
            else
            {
                result = _Service.ToReport(ReportNo, IsToPDF, out FileName);
            }

            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }


        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult SendMail(string ReportNo)
        {
            this.CheckSession();

            BaseResult result = null;
            string FileName = string.Empty;

            result = _Service.ToReport(ReportNo, true, out FileName);

            if (!result.Result)
            {
                result.ErrorMessage = result.ErrorMessage.ToString();
            }
            string reportPath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + FileName;

            return Json(new { Result = result.Result, ErrorMessage = result.ErrorMessage, reportPath = reportPath, FileName = FileName });
        }
    }
}