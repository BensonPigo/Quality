using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
using Ict.Data.Defs;
using NPOI.SS.Formula.Functions;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;
using static Quality.Helper.Attribute;
using Orders = DatabaseObject.ProductionDB.Orders;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class DailyMoistureController : BaseController
    {
        private DailyMoistureService _Service;
        public DailyMoistureController()
        {
            _Service = new DailyMoistureService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.DailyMoisture,,";
        }
        // GET: BulkFGT/DailyMoisture
        public ActionResult Index()
        {
            DailyMoisture_ViewModel model = new DailyMoisture_ViewModel()
            {
                Main = new DailyMoisture_Result(),
                Request = new DailyMoisture_Request(),
                Details = new List<DailyMoisture_Detail_Result>(),
                EndlineMoisture_Source = _Service.GetEndlineMoisture(),
                Action_Source = _Service.GetAction(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string OrderID)
        {
            DailyMoisture_Request Req = new DailyMoisture_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                OrderID = OrderID,
            };
            DailyMoisture_ViewModel model = new DailyMoisture_ViewModel()
            {
                Request = Req,
                Main = new DailyMoisture_Result(),
                Details = new List<DailyMoisture_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            model = _Service.GetDailyMoisture(Req);
            model.EndlineMoisture_Source = _Service.GetEndlineMoisture();
            model.Action_Source = _Service.GetAction();

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
        public ActionResult Query(DailyMoisture_ViewModel Req)
        {
            DailyMoisture_ViewModel model = new DailyMoisture_ViewModel()
            {
                Request = Req.Request,
                Main = new DailyMoisture_Result(),
                Details = new List<DailyMoisture_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };

            model = _Service.GetDailyMoisture(Req.Request);
            model.EndlineMoisture_Source = _Service.GetEndlineMoisture();
            model.Action_Source = _Service.GetAction();

            if (model.Result && !model.ReportNo_Source.Any())
            {
                model.ErrorMessage = $@"msg.WithInfo('Data not found.')";
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
        public ActionResult NewSave(DailyMoisture_ViewModel Req)
        {
            DailyMoisture_ViewModel model = new DailyMoisture_ViewModel()
            {
                Main = new DailyMoisture_Result(),
                Request = new DailyMoisture_Request(),
                Details = new List<DailyMoisture_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };


            if (Req.Details == null)
            {
                Req.Details = new List<DailyMoisture_Detail_Result>();
            }


            Req.Main.AddName = this.UserID;

            BaseResult result = _Service.Create(Req, this.MDivisionID, this.UserID, out string NewReportNo);

            if (!result.Result)
            {
                string ErrorMessage = result.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
            }
            else if (result.Result)
            {
                model = _Service.GetDailyMoisture(new DailyMoisture_Request()
                {
                    ReportNo = NewReportNo,
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    OrderID = Req.Main.OrderID,
                });
            }
            model.EndlineMoisture_Source = _Service.GetEndlineMoisture();
            model.Action_Source = _Service.GetAction();

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult EditSave(DailyMoisture_ViewModel Req)
        {
            DailyMoisture_ViewModel model = new DailyMoisture_ViewModel()
            {
                Main = new DailyMoisture_Result(),
                Request = new DailyMoisture_Request(),
                Details = new List<DailyMoisture_Detail_Result>(),
                ReportNo_Source = new List<SelectListItem>(),
            };


            if (Req.Details == null)
            {
                Req.Details = new List<DailyMoisture_Detail_Result>();
            }

            Req.Main.AddName = this.UserID;

            BaseResult result = _Service.Update(Req, this.UserID);

            if (!result.Result)
            {
                string ErrorMessage = result.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
            }
            else if (result.Result)
            {
                model = _Service.GetDailyMoisture(new DailyMoisture_Request()
                {
                    ReportNo = Req.Main.ReportNo,
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    OrderID = Req.Main.OrderID,
                });
            }
            model.EndlineMoisture_Source = _Service.GetEndlineMoisture();
            model.Action_Source = _Service.GetAction();

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(DailyMoisture_ViewModel Req)
        {

            DailyMoisture_ViewModel model = new DailyMoisture_ViewModel()
            {
                Main = new DailyMoisture_Result(),
                Request = Req.Request,
                Details = new List<DailyMoisture_Detail_Result>(),
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
                model = _Service.GetDailyMoisture(new DailyMoisture_Request()
                {
                    BrandID = Req.Request.BrandID,
                    SeasonID = Req.Request.SeasonID,
                    StyleID = Req.Request.StyleID,
                    OrderID = Req.Request.OrderID,
                });
            }
            model.EndlineMoisture_Source = _Service.GetEndlineMoisture();
            model.Action_Source = _Service.GetAction();

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Encode")]
        public ActionResult Encode(DailyMoisture_ViewModel Req)
        {
            DailyMoisture_ViewModel model = new DailyMoisture_ViewModel();
            Req.Main = new DailyMoisture_Result()
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
                model = _Service.GetDailyMoisture(new DailyMoisture_Request()
                {
                    ReportNo = Req.Request.ReportNo,
                    BrandID = Req.Request.BrandID,
                    SeasonID = Req.Request.SeasonID,
                    StyleID = Req.Request.StyleID,
                    OrderID = Req.Request.OrderID,
                });

                if (model.Main.Result == "Fail")
                {
                    model.ErrorMessage = $@"FailMail();";
                }
            }
            model.EndlineMoisture_Source = _Service.GetEndlineMoisture();
            model.Action_Source = _Service.GetAction();

            return View("Index", model);
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Amend")]
        public ActionResult Amend(DailyMoisture_ViewModel Req)
        {
            DailyMoisture_ViewModel model = new DailyMoisture_ViewModel();
            Req.Main = new DailyMoisture_Result()
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
                model = _Service.GetDailyMoisture(new DailyMoisture_Request()
                {
                    ReportNo = Req.Request.ReportNo,
                    BrandID = Req.Request.BrandID,
                    SeasonID = Req.Request.SeasonID,
                    StyleID = Req.Request.StyleID,
                    OrderID = Req.Request.OrderID,
                });
            }
            model.EndlineMoisture_Source = _Service.GetEndlineMoisture();
            model.Action_Source = _Service.GetAction();

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
                //List<Order_Qty> order_qtyResult = _Service.GetDistinctArticle(order_qty);

                //if (order_qtyResult.Any())
                //{
                //    Article = order_qtyResult.FirstOrDefault().Article;
                //}

            }

            return Json(new { ErrMsg = "", BrandID = BrandID, SeasonID = SeasonID, StyleID = StyleID,  });
        }


        [SessionAuthorizeAttribute]
        [HttpPost]
        public ActionResult AddDetailRow(int lastNO, string BrandID, string SeasonID, string StyleID)
        {
            List<SelectListItem> areaSelect = new List<SelectListItem>()
                {
                    new SelectListItem(){ Text="Main fabric",Value="Main fabric"},
                    new SelectListItem(){ Text="Pocket seams",Value="Pocket seams"},
                    new SelectListItem(){ Text="Side seam (applicable if garment without pocket)",Value="Side seam (applicable if garment without pocket)"},
                    new SelectListItem(){ Text="Inseam (applicable if garment without pocket)",Value="Inseam (applicable if garment without pocket)"},
                    new SelectListItem(){ Text="Let hem/ bottom hem",Value="Let hem/ bottom hem"},
                    new SelectListItem(){ Text="Neckline/ Collar (back)",Value="Neckline/ Collar (back)"},
                    new SelectListItem(){ Text="Printing (if have)",Value="Printing (if have)"},
                    new SelectListItem(){ Text="Hood seam",Value="Hood seam"},
                    new SelectListItem(){ Text="Elastic (front)",Value="Elastic (front)"},
                    new SelectListItem(){ Text="Elastic (back)",Value="Elastic (back)"},
                    new SelectListItem(){ Text="Elastic (inner)",Value="Elastic (inner)"},
                };
            List<SelectListItem> fabricSelect = new List<SelectListItem>()
            {
                new SelectListItem(){Text="A",Value="A"},
                new SelectListItem(){Text="B",Value="B"},
                new SelectListItem(){Text="C",Value="C"},
                new SelectListItem(){Text="D",Value="D"},
            };

            List<SelectListItem> resultSelect = new List<SelectListItem>()
            {
                new SelectListItem(){Text="Pass",Value="Pass"},
                new SelectListItem(){Text="Fail",Value="Fail"},
            };


            MockupOven_ViewModel model = new MockupOven_ViewModel();

            string html = "";
            html += $"<tr idx='{lastNO}'>";

            html += $"<td> <input id='Seq{lastNO}' idx='{lastNO}' type ='hidden'> ";
            html += $"<input id='Details_{lastNO}__Point1' name='Details[{lastNO}].Point1' class='OnlyEdit' type='number' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)' style='width:85%;'>";
            html += $"</td>";


            html += $"<td>";
            html += $"<input id='Details_{lastNO}__Point2' name='Details[{lastNO}].Point2' class='OnlyEdit' type='number' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)' style='width:85%;'>";
            html += $"</td>";

            html += $"<td>";
            html += $"<input id='Details_{lastNO}__Point3' name='Details[{lastNO}].Point3' class='OnlyEdit' type='number' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)' style='width:85%;'>";
            html += $"</td>";

            html += $"<td>";
            html += $"<input id='Details_{lastNO}__Point4' name='Details[{lastNO}].Point4' class='OnlyEdit' type='number' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)' style='width:85%;'>";
            html += $"</td>";

            html += $"<td>";
            html += $"<input id='Details_{lastNO}__Point5' name='Details[{lastNO}].Point5' class='OnlyEdit' type='number' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)' style='width:85%;'>";
            html += $"</td>";

            html += $"<td>";
            html += $"<input id='Details_{lastNO}__Point6' name='Details[{lastNO}].Point6' class='OnlyEdit' type='number' step='0.01' type='number' value='0' onchange='value=QtyCheck(value)' style='width:85%;'>";
            html += $"</td>";

            html += $"<td><select id='Details_{lastNO}__Area' name='Details[{lastNO}].Area'  class='OnlyEdit' style='width:90%' >";
            foreach (var val in areaSelect)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";


            html += $"<td><select id='Details_{lastNO}__Fabric' name='Details[{lastNO}].Fabric'  class='OnlyEdit' >";
            foreach (var val in fabricSelect)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";


            html += $"<td></td>";

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
        public JsonResult SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = _Service.SendMail(ReportNo, TO, CC, Subject,Body,Files);
            return Json(result);
        }
    }
}