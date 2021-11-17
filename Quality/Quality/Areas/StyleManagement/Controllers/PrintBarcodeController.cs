using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service.StyleManagement;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.SampleRFT;
using Quality.Controllers;
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

            List<PrintBarcode_ViewModel> model = new List<PrintBarcode_ViewModel>();
            return View(model);
        }

        [HttpPost]
        public ActionResult Query(StyleManagement_Request Req)
        {
            this.CheckSession();

            List<PrintBarcode_ViewModel> model = new List<PrintBarcode_ViewModel>();

            try
            {
                bool st = string.IsNullOrEmpty(Req.StyleID);
                bool sb = (string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID));
                if (st && sb)
                {
                    ViewData["MsgScript"] = $@"msg.WithInfo(""Please input ＜Style＞ or ＜Brand, Season＞."");";
                    return View("index", model);
                }

                model = _Service.Get_PrintBarcodeStyleInfo(Req).ToList();

                if (!model.Any())
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
        public ActionResult PrintBarcode(List<int> DataList)
        {
            this.CheckSession();

            TempData["StyleUkeyList"] = DataList;
            return Json(true);
        }
    }
}