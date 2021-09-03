using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service.StyleManagement;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
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
        private IStyleManagementService _StyleManagementService;

        public PrintBarcodeController()
        {
            _StyleManagementService = new StyleManagementService();
        }

        // GET: StyleManagement/PrintBarcode
        public ActionResult Index()
        {
            this.CheckSession();

            PrintBarcode_ViewModel model = new PrintBarcode_ViewModel()
            {
                DataList = new List<StyleResult_ViewModel>()
            };
            return View(model);
        }

        [HttpPost]
        public ActionResult Query(StyleManagement_Request Req)
        {
            this.CheckSession();

            PrintBarcode_ViewModel model = new PrintBarcode_ViewModel()
            {
                DataList = new List<StyleResult_ViewModel>()
            };
            try
            {
                bool st = string.IsNullOrEmpty(Req.StyleID);
                bool sb = (string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID));
                if (st && sb)
                {
                    model.MsgScript = $@"
msg.WithInfo('Please input ＜Style＞ or ＜Brand, Season＞.');
";
                    return View("index", model);
                }

                model.DataList = _StyleManagementService.Get_PrintBarcodeStyleInfo(Req).ToList();

                if (!model.DataList.Any())
                {
                    model.MsgScript = $@"
msg.WithInfo('Data not found.');
";
                }
            }
            catch (Exception ex)
            {
                model.MsgScript = $@"
msg.WithInfo('Error : {ex.Message}');
";
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