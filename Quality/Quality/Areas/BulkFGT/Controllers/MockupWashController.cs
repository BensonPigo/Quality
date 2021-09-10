using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
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
    public class MockupWashController : BaseController
    {
        private IMockupWashService _MockupWashService;
        public MockupWashController()
        {
            _MockupWashService = new MockupWashService();
        }

        // GET: BulkFGT/MockupWash
        public ActionResult Index()
        {

            MockupWash_ViewModel model = new MockupWash_ViewModel()
            {
                MockupWash_Detail = new List<MockupWash_Detail_ViewModel>(),
                ReportNo_Source = new List<string>(),
                Request = new MockupWash_Request(),
                TestingMethod_Source = new List<SelectListItem>()
            };

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source;
            ViewBag.ArtworkTypeID_Source = new SetListItem().ItemListBinding(new List<string>());
            ViewBag.AccessoryRefNo_Source = new SetListItem().ItemListBinding(new List<string>());

            return View(model);
        }


        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(MockupWash_ViewModel Req)
        {

            MockupWash_Request MockupWash = new MockupWash_Request()
            { BrandID = Req.Request.BrandID, SeasonID = Req.Request.SeasonID, StyleID = Req.Request.StyleID };

            var model = _MockupWashService.GetMockupWash(MockupWash);
            model.Request = new MockupWash_Request();
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(Req.Request.BrandID, Req.Request.SeasonID, Req.Request.StyleID);
            return View("Index", model);
        }


        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult NewSave(MockupWash_ViewModel Req)
        {

            MockupWash_Request MockupWash = new MockupWash_Request()
            { BrandID = Req.BrandID, SeasonID = Req.SeasonID, StyleID = Req.StyleID };

            var model = _MockupWashService.GetMockupWash(MockupWash);
            model.Request = new MockupWash_Request();
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.BrandID, Req.SeasonID, Req.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(Req.BrandID, Req.SeasonID, Req.StyleID);
            return View("Index", model);
        }


        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult EditSave(MockupWash_ViewModel Req)
        {
            MockupWash_Request MockupWash = new MockupWash_Request()
            { BrandID = Req.BrandID, SeasonID = Req.SeasonID, StyleID = Req.StyleID };

            var model = _MockupWashService.GetMockupWash(MockupWash);
            model.Request = new MockupWash_Request();
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.BrandID, Req.SeasonID, Req.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(Req.BrandID, Req.SeasonID, Req.StyleID);
            return View("Index", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult DeleteReportNo(MockupWash_ViewModel Req)
        {
            MockupWash_Request MockupWash = new MockupWash_Request()
            { BrandID = Req.BrandID, SeasonID = Req.SeasonID, StyleID = Req.StyleID };

            var model = _MockupWashService.GetMockupWash(MockupWash);
            model.Request = new MockupWash_Request();
            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(model.ReportNo_Source);
            ViewBag.ResultList = model.Result_Source; ;
            ViewBag.ArtworkTypeID_Source = GetArtworkTypeIDList(Req.BrandID, Req.SeasonID, Req.StyleID);
            ViewBag.AccessoryRefNo_Source = GetAccessoryRefNoList(Req.BrandID, Req.SeasonID, Req.StyleID);
            return View("Index", model);
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

            return RedirectToAction("Index");
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
            List<Orders> orderResult = _MockupWashService.GetOrders(order);

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
                List<Order_Qty> order_qtyResult = _MockupWashService.GetDistinctArticle(order_qty);

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
            return _MockupWashService.GetArtworkTypeID(styleArtwork_Request);
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
            return _MockupWashService.GetAccessoryRefNo(AccessoryRefNo_Request);
        }


        [HttpPost]
        public ActionResult AddDetailRow(int lastNO, string BrandID, string SeasonID, string StyleID)
        {
            List<SelectListItem> AccessoryRefNo_Source = new SetListItem().ItemListBinding(new List<string>());

            if (!string.IsNullOrWhiteSpace(BrandID) && !string.IsNullOrWhiteSpace(SeasonID) && !string.IsNullOrWhiteSpace(StyleID))
            {
                AccessoryRefNo_Source = GetAccessoryRefNoList(BrandID, SeasonID, StyleID);
            }


            MockupWash_ViewModel model = new MockupWash_ViewModel();

            int i = lastNO;
            string html = "";
            html += "<tr>";
            html += "<td><input id='Seq' idx=" + i + " type ='hidden'></input> <input id='MockupWash_Detail_" + i + "__TypeofPrint' name='MockupWash_Detail[" + i + "].TypeofPrint' class='OnlyEdit' type='text' value=''></td>";
            html += "<td><input id='MockupWash_Detail_" + i + "__Design' name='MockupWash_Detail[" + i + "].Design' class='OnlyEdit' type='text' ></td>";
            html += "<td><div class='input-group'><input id='MockupWash_Detail_" + i + "__ArtworkColor' name='MockupWash_Detail[" + i + "].ArtworkColor'  class ='AFColor' type='hidden'><input id='MockupWash_Detail_" + i + "__ArtworkColorName' name='MockupWash_Detail[" + i + "].ArtworkColorName'  class ='AFColor' type='text' readonly='readonly'> <input  idv='" + i.ToString() + "' type='button' class='btnArtworkColorItem  site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>";
            html += "<td><select id='MockupWash_Detail_" + i + "__AccessoryRefNo_Source' name='MockupWash_Detail[" + i + "].AccessoryRefNo_Source'  class='OnlyEdit' style='width: 157px;'><option value=''></option>";
            foreach (var val in AccessoryRefNo_Source)
            {
                html += "<option value='" + val.Value + "'>" + val.Text + "</option>";
            }
            html += "</select></td>";
            html += "<td><input id='MockupWash_Detail_" + i + "__FabricRefNo' name='MockupWash_Detail[" + i + "].FabricRefNo' type='text' ></td>";
            html += "<td><div class='input-group'><input id='MockupWash_Detail_" + i + "__FabricColor' name='MockupWash_Detail[" + i + "].FabricColor'  class ='AFColor' type='hidden'><input id='MockupWash_Detail_" + i + "__FabricColorName' name='MockupWash_Detail[" + i + "].FabricColorName'  class ='AFColor' type='text' readonly='readonly'> <input  idv='" + i.ToString() + "' type='button' class='btnFabricColorItem  site-btn btn-blue' style='margin: 0; border: 0; ' value='...' /></div></td>";

            html += "<td><select  id='MockupWash_Detail_" + i + "__Result' name='MockupWash_Detail[" + i + "].Result' class='OnlyEdit result' onchange='changeResult()' style='width: 157px;' ><option value=''></option>";
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

            html += "<td><input id='MockupWash_Detail_" + i + "__SCIRefno' name='MockupWash_Detail[" + i + "].Remark' type='text' class='OnlyEdit'></td>";
            html += "<td><input id='MockupWash_Detail_" + i + "__ColorID' name='MockupWash_Detail[" + i + "].LastUpdate' type='text'  readonly='readonly' ></td>";

            html += "<td> <div style='width: 5vw;'><img  class='detailDelete' src='/Image/Icon/Delete.png' width='30'> </div></td>";
            html += "</tr>";

            return Content(html);
        }
    }
}