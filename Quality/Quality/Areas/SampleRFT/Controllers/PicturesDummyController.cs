
using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using Ionic.Zip;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.Drawing;
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

        public ActionResult IndexGet(string OrderID)
        {
            this.CheckSession();

            if (string.IsNullOrEmpty(OrderID))
            {
                RFT_PicDuringDummyFitting_ViewModel e = new RFT_PicDuringDummyFitting_ViewModel()
                {
                    ErrorMessage = "SP# cannot be empty",
                    DataList = new List<RFT_PicDuringDummyFitting>()
                };
                return View("Index", e);
            }

            RFT_PicDuringDummyFitting_ViewModel model = _PicturesDummyService.Get_PicturesDummy_Result(new RFT_PicDuringDummyFitting_ViewModel()
            {
                OrderID = OrderID,
                OrderTypeID = "OrderID"
            }
                );

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo('{model.ErrorMessage.Replace("\r\n", "<br />")}');";
            }

            return View("Index", model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
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
            model.OrderID = Req.OrderID;
            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo('{model.ErrorMessage.Replace("\r\n", "<br />")}');";
            }

            return View("Index", model);
        }


        public ActionResult DownloadPictures(RFT_PicDuringDummyFitting Req)
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

            RFT_PicDuringDummyFitting_ViewModel model = _PicturesDummyService.Get_PicturesDummy_Result(new RFT_PicDuringDummyFitting_ViewModel() 
            { 
                OrderID = Req.OrderID
            });

            model.OrderID = Req.OrderID;

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo('{model.ErrorMessage.Replace("\r\n", "<br />")}');";
                return View("Index", model);
            }

            RFT_PicDuringDummyFitting pic = model.DataList.Where(o => o.Article == Req.Article && o.Size == Req.Size).FirstOrDefault();


            ZipFile zip = new ZipFile();
            if (pic.Front!=null)
            {
                MemoryStream msA = new MemoryStream(pic.Front);
                Image imgFront = Image.FromStream(msA);
                string frontName = $"{Req.OrderID}_{pic.Article}_{pic.Size}_Front.png";
                string frontPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", frontName);
                imgFront.Save(frontPath);
                zip.AddFile(frontPath, string.Empty);
            }
            if (pic.Side != null)
            {
                MemoryStream msB = new MemoryStream(pic.Side);
                Image imgSide = Image.FromStream(msB);
                string sideName = $"{Req.OrderID}_{pic.Article}_{pic.Size}_Side.png";
                string sidePath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", sideName);
                imgSide.Save(sidePath);
                zip.AddFile(sidePath, string.Empty);
            }
            if (pic.Back != null)
            {
                MemoryStream msC = new MemoryStream(pic.Back);
                Image imgBack = Image.FromStream(msC);
                string BackName = $"{Req.OrderID}_{pic.Article}_{pic.Size}_Back.png";
                string backPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", BackName);
                imgBack.Save(backPath);
                zip.AddFile(backPath, string.Empty);

            }
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", $"attachment; filename=Dummy Fitting_{DateTime.Now.ToString("yyyyMMddHHmmss")}.zip");
            zip.Save(Response.OutputStream);

            Response.Flush();
            Response.End();

            return View("Index", model);
        }


        [HttpPost]
        [SessionAuthorizeAttribute]
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