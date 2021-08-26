
using BusinessLogicLayer.Interface.SampleRFT;
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
    public class CFTCommentsController : BaseController
    {
        private ICFTCommentsService _ICFTCommentsService;

        public CFTCommentsController()
        {
            _ICFTCommentsService = new CFTCommentsService();
            this.SelectedMenu = "Sample RFT";

        }


        // GET: SampleRFT/CFTComments
        public ActionResult Index()
        {
            this.CheckSession();
            CFTComments_ViewModel model = new CFTComments_ViewModel() { QueryType = "Style",DataList = new List<CFTComments_Result>()};

            if (TempData["Model"] != null)
            {
                model =(CFTComments_ViewModel)TempData["Model"];
            }

            return View(model);
        }

        [HttpPost]
        public ActionResult Query(CFTComments_ViewModel Req)
        {
            this.CheckSession();

            CFTComments_ViewModel model = new CFTComments_ViewModel();
            Req.DataList = new List<CFTComments_Result>();

            if (Req.QueryType == "OrderID")
            {
                if (Req.OrderID == null || string.IsNullOrEmpty(Req.OrderID))
                {

                    Req.ErrorMessage = $@"
msg.WithInfo('SP# cannot be emptry');
";
                    return View("Index", Req);
                }

                model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel() { OrderID = Req.OrderID });

                if (model.OrderID == null)
                {
                    Req.ErrorMessage = $@"
msg.WithInfo('Cannot found SP# {Req.OrderID}');
";
                    return View("Index", Req);
                }

            }
            else if (Req.QueryType == "Style")
            {
                if (string.IsNullOrEmpty(Req.StyleID) || string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID)
                    || Req.StyleID == null || Req.BrandID == null || Req.SeasonID == null)
                {

                    Req.ErrorMessage = $@"
msg.WithInfo('Style#, Brand and Season cannot be emptry');
";
                    return View("Index", Req);
                }

                model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel()
                {
                    StyleID = Req.StyleID,
                    BrandID = Req.BrandID,
                    SeasonID = Req.SeasonID,
                });

                if (model.OrderID == null)
                {
                    Req.ErrorMessage = $@"
msg.WithInfo('Cannot found combination Style# {Req.StyleID}, Brand {Req.BrandID}, Season {Req.SeasonID}');
";
                    return View("Index", Req);
                }
            }

            //fake data
            model = _ICFTCommentsService.Get_CFT_OrderComments(model);
            //model.DataList = new List<CFTComments_Result>() {

            //    new CFTComments_Result(){ SampleStage ="111",CommentsCategory="222",Comnments="333"},
            //    new CFTComments_Result(){ SampleStage ="111",CommentsCategory="222",Comnments="333"},
            //    new CFTComments_Result(){ SampleStage ="111",CommentsCategory="222",Comnments="333"},
            //    new CFTComments_Result(){ SampleStage ="111",CommentsCategory="222",Comnments="333"},
            //};

            if (!model.Result)
            {
                model.ErrorMessage = $@"
msg.WithInfo('{model.ErrorMessage.Replace("'",string.Empty)}');
";
            }
            model.QueryType = Req.QueryType;
            TempData["Model"] = model;

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult GetOrderinfo(string OrderID)
        {
            this.CheckSession();
            CFTComments_ViewModel model = _ICFTCommentsService.Get_CFT_Orders(new CFTComments_ViewModel() { OrderID=OrderID});

            return View();
        }
    }
}