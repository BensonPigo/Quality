
using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;


namespace Quality.Areas.SampleRFT.Controllers
{
    public class PicturesDummyController : BaseController
    {
        private IPicturesDummyService _PicturesDummyService;

        public PicturesDummyController()
        {
            _PicturesDummyService = new PicturesDummyService();
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.PicturesDummy,,";
        }

        // GET: SampleRFT/PicturesDummy
        public ActionResult Index()
        {
            this.CheckSession();
            RFT_PicDuringDummyFitting_ViewModel model = new RFT_PicDuringDummyFitting_ViewModel();
            model.DataList = new List<RFT_PicDuringDummyFitting>();
            return View(model);
        }

        [HttpPost]
        public ActionResult Query(RFT_PicDuringDummyFitting_ViewModel Req)
        {
            this.CheckSession();


            if (Req == null || string.IsNullOrEmpty(Req.OrderID))
            {
                RFT_PicDuringDummyFitting_ViewModel e = new RFT_PicDuringDummyFitting_ViewModel()
                {
                    ErrorMessage = "SP# cannot be empty",
                    DataList = new List<RFT_PicDuringDummyFitting>()
                };
                return View("Index", e);
            }

            RFT_PicDuringDummyFitting_ViewModel model = _PicturesDummyService.Get_PicturesDummy_Result(Req);

            if (!model.Result)
            {
                model.ErrorMessage = $@"
msg.WithInfo('{model.ErrorMessage.Replace("\r\n", "<br />")}');
";
            }

            return View("Index", model);
        }


        [HttpPost]
        public ActionResult CheckOrder(string OrderID)
        {
            this.CheckSession();

            try
            {
                RFT_PicDuringDummyFitting_ViewModel model = _PicturesDummyService.Check_OrderID_Exists(OrderID);

                return Json(model);
            }
            catch (Exception ex)
            {
                return Json(ex.Message);
            }
        }
    }
}