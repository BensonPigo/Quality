using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service.StyleManagement;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.SampleRFT;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.StyleManagement
{
    public class PrintBarcodeController : BaseController
    {
        private PrintBarcodeService _Service;

        public PrintBarcodeController()
        {
            _Service = new PrintBarcodeService();
            this.SelectedMenu = "Style Management";
            ViewBag.OnlineHelp = this.OnlineHelp + "StyleManagement.PrintBarcode,,";
        }

        // GET: StyleManagement/PrintBarcode
        public ActionResult Index()
        {
            this.CheckSession();

            PrintBarcode_ViewModel model = new PrintBarcode_ViewModel()
            {
                Details = new List<PrintBarcode_Detail>(),
                SampleStage_Source = new List<SelectListItem>()
            };
            return View(model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult Query(StyleManagement_Request Req)
        {
            this.CheckSession();

            PrintBarcode_ViewModel model = new PrintBarcode_ViewModel()
            {
                Details = new List<PrintBarcode_Detail>()
            };

            try
            {
                bool st = string.IsNullOrEmpty(Req.StyleID);
                bool sb = (string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID));
                if (st && sb)
                {
                    ViewData["MsgScript"] = $@"msg.WithInfo(""Please input ＜Style＞ or ＜Brand, Season＞."");";
                    return View("index", model);
                }

                model.Details = _Service.Get_PrintBarcodeStyleInfo(Req).ToList();
                model.SampleStage_Source = _Service.Get_SampleStage(Req).ToList();

                if (!model.Details.Any())
                {
                    ViewData["MsgScript"] = $@"msg.WithInfo(""Data not found."");";
                }
            }
            catch (Exception ex)
            {
                ViewData["MsgScript"] = $@"msg.WithInfo(""Error : {ex.Message.Replace("\r\n", "<br />")}"");";
            }

            return View("index", model);
        }

        public ActionResult PrintBarcode()
        {
            this.CheckSession();

            ViewData["StyleUkeyList"] = TempData["StyleUkeyList"];
            return View();
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult PrintBarcode(List<PrintBarcode_Detail> DataList)
        {
            this.CheckSession();

            TempData["StyleUkeyList"] = DataList;
            return Json(true);
        }
    }
}