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
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;
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
                ArtworkType_Source = new List<SelectListItem>(),
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
                model.ErrorMessage = "msg.WithInfo('Data not found')";
            }
            else if (!model.Result)
            {
                string ErrorMessage = model.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
                model.ReportNo_Source = new List<SelectListItem>();
            }

            var tmpArtworkType_Source = _Service.GetArtworkTypeList(new Orders()
            {
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
            });

            model.ArtworkType_Source = tmpArtworkType_Source.Any() ? tmpArtworkType_Source : new List<SelectListItem>();

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
                model.ErrorMessage = "msg.WithInfo('Data not found')";
            }
            else if (!model.Result)
            {
                string ErrorMessage = model.ErrorMessage;
                model.ErrorMessage = $@"msg.WithInfo(""{ErrorMessage}"")";
                model.ReportNo_Source = new List<SelectListItem>();
            }

            var tmpArtworkType_Source = _Service.GetArtworkTypeList(new Orders()
            {
                BrandID = Req.Request.BrandID,
                SeasonID = Req.Request.SeasonID,
                StyleID = Req.Request.StyleID,
            });

            model.ArtworkType_Source = tmpArtworkType_Source.Any() ? tmpArtworkType_Source : new List<SelectListItem>();
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
            var tmpArtworkType_Source = _Service.GetArtworkTypeList(new Orders()
            {
                BrandID = Req.Main.BrandID,
                SeasonID = Req.Main.SeasonID,
                StyleID = Req.Main.StyleID,
            });

            model.ArtworkType_Source = tmpArtworkType_Source.Any() ? tmpArtworkType_Source : new List<SelectListItem>();

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

            var tmpArtworkType_Source = _Service.GetArtworkTypeList(new Orders()
            {
                BrandID = Req.Request.BrandID,
                SeasonID = Req.Request.SeasonID,
                StyleID = Req.Request.StyleID,
            });

            model.ArtworkType_Source = tmpArtworkType_Source.Any() ? tmpArtworkType_Source : new List<SelectListItem>();
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

            var tmpArtworkType_Source = _Service.GetArtworkTypeList(new Orders()
            {
                BrandID = Req.Request.BrandID,
                SeasonID = Req.Request.SeasonID,
                StyleID = Req.Request.StyleID,
            });

            model.ArtworkType_Source = tmpArtworkType_Source.Any() ? tmpArtworkType_Source : new List<SelectListItem>();
            return View("Index", model);
        }

        public JsonResult Encode(string ReportNo)
        {
            CheckSession();
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel();
            HeatTransferWash_ViewModel Req = new HeatTransferWash_ViewModel() 
            { 
             Main = new HeatTransferWash_Result()
             {
                 ReportNo = ReportNo,
                 Status = "Confirmed",
                 EditName = this.UserID,
             }
            };
            BaseResult result = _Service.EncodeAmend(Req);

            model = _Service.GetHeatTransferWash(new HeatTransferWash_Request() { ReportNo = ReportNo });

            if (model.Main.Result == "Fail")
            {
                return Json(new { result.Result, ErrMsg = result.ErrorMessage, Action = "FailMail()" });
            }
            else
            {
                return Json(new { result.Result, ErrMsg = result.ErrorMessage, Action = "" });
            }
        }

        public ActionResult Amend(string ReportNo)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel();
            HeatTransferWash_ViewModel Req = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result() 
                { 
                    ReportNo = ReportNo,
                    Status = "New",
                    EditName = this.UserID, 
                }
            };
            BaseResult result = _Service.EncodeAmend(Req);
            

            return Json(new { result.Result, ErrMsg = result.ErrorMessage });
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
            List<SelectListItem> ArtworkTypeList = new List<SelectListItem>();
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
                ArtworkTypeList = _Service.GetArtworkTypeList(new Orders()
                {
                    BrandID = BrandID,
                    SeasonID = SeasonID,
                    StyleID = StyleID,
                });

                if (order_qtyResult.Any())
                {
                    Article = order_qtyResult.FirstOrDefault().Article;
                }

            }

            return Json(new { ErrMsg = "", BrandID = BrandID, SeasonID = SeasonID, StyleID = StyleID, Article = Article, Teamwear = Teamwear, ArtworkTypeList = ArtworkTypeList });
        }

        [SessionAuthorizeAttribute]
        [HttpPost]
        public JsonResult GetLastDetailData(string HTRefNo)
        {
            string errorMsg = string.Empty;

            HeatTransferWash_Detail_Result detail = new HeatTransferWash_Detail_Result();
            try
            {
                detail = _Service.GetLastDetailData(HTRefNo);
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
            }



            return Json(new { ErrMsg = errorMsg, DetailData = detail});
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


            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel();

            string html = "";
            html += $"<tr idx='{lastNO}'>";

            html += $"<td> <input id='Seq{lastNO}' idx='{lastNO}' type ='hidden'> ";
            html += $"<input id='Details_{lastNO}__FabricRefNo' name='Details[{lastNO}].FabricRefNo' class='OnlyEdit FabricRefNoTxt' type='text' value=''  style='width:80%;'>";
            html += $"<input targetID='Details_{lastNO}__FabricRefNo' type='button' class='btnRefnoSelectItem site-btn btn-blue' style='margin:0;border:0;' value='...' />";
            html += $"</td>";


            html += $"<td> <input id='Details_{lastNO}__HTRefNo' name='Details[{lastNO}].HTRefNo' class='OnlyEdit HTRefNoTxt' type='text' value=''  style='width:80%;'>";
            html += $"     <input targetID='Details_{lastNO}__HTRefNo' type='button' class='btnHTRefNoSelectItem site-btn btn-blue' style='margin:0;border:0;' value='...' />";
            html += $"</td>";

            html += $"<td><input id='Details_{lastNO}__Temperature' name='Details[{lastNO}].Temperature' type='number' style='width:100%;' class='OnlyEdit'></td>";
            html += $"<td><input id='Details_{lastNO}__Time' name='Details[{lastNO}].Time' type='number' style='width:100%;' class='OnlyEdit'></td>";
            html += $"<td><input id='Details_{lastNO}__SecondTime' name='Details[{lastNO}].SecondTime' type='number' style='width:100%;' class='OnlyEdit'></td>";
            html += $"<td><input id='Details_{lastNO}__Pressure' name='Details[{lastNO}].Pressure' type='number' step='0.01' max='100' onchange = 'value=PressureCheck(value)' style='width:100%;' class='OnlyEdit'></td>";
            html += $"<td><input id='Details_{lastNO}__PeelOff' name='Details[{lastNO}].PeelOff' type='text' maxlength='5' style='width:100%;' class='OnlyEdit'></td>";

            html += $"<td><select id='Details_{lastNO}__Cycles' name='Details[{lastNO}].Cycles'  class='OnlyEdit' >";
            foreach (var item in model.Cycles_Source)
            {
                html += "<option value='" + item.Value + "'>" + item.Text + "</option>";
            }
            html += "</select></td>";

            if (BrandID == "ADIDAS")
            {
                html += $"<td><select id='Details_{lastNO}__TemperatureUnit' name='Details[{lastNO}].TemperatureUnit'  class='OnlyEdit' >";
                foreach (var item in model.TemperatureUnit_Source)
                {
                    html += "<option value='" + item.Value + "'>" + item.Text + "</option>";
                }
                html += "</select></td>";
            }
            else
            {
                html += $"<td><input id='Details_{lastNO}__TemperatureUnit' name='Details[{lastNO}].TemperatureUnit' onchange = 'value=WashingIntCheck(value)' type='number' style='width:100%;' class='OnlyEdit'></td>";
            }


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
        public ActionResult Report(string ReportNo, bool IsToPDF)
        {
            BaseResult result = _Service.ToReport(ReportNo, IsToPDF, out string FileName);
            string reportPath = "/TMP/" + Uri.EscapeDataString(FileName);
            return Json(new { result.Result, result.ErrorMessage, reportPath });
        }


        [HttpPost]
        [SessionAuthorizeAttribute]
        public JsonResult SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = _Service.SendMail(ReportNo, TO, CC, Subject, Body, Files);
            return Json(result);
        }
    }
}