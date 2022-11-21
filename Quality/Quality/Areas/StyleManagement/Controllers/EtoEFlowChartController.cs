using BusinessLogicLayer.Service.StyleManagement;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.StyleManagement;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DatabaseObject.ResultModel.EtoEFlowChart;

namespace Quality.Areas.StyleManagement.Controllers
{
    public class EtoEFlowChartController : BaseController
    {
        public EtoEFlowChartService _Service;

        public EtoEFlowChartController()
        {
            _Service = new EtoEFlowChartService();
            this.SelectedMenu = "Style Management";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.EtoEFlowChart,,";
        }

        // GET: StyleManagement/EtoEFlowChart
        public ActionResult Index()
        {
            EtoEFlowChart_ViewModel model = new EtoEFlowChart_ViewModel()
            {
                Request = new EtoEFlowChart_Request(),
                Warehouse = new Warehouse()
                { 
                    PhysicalInspection_Detail = new List<Warehouse_PhysicalInspection>(),
                },
                Development = new Development()
                {
                    SampleRFT_Detail = new List<Development_SampleRFT>(),
                    TestResult_Detail = new List<Development_TestResult>(),
                    RRLR_Detail = new List<Development_RRLR>(),
                    FD_Detail = new List<Development_FD>(),
                    CFTComments = new List<Development_SampleRFT_CFTComments>()
                },
                Production = new Production()
                { 
                    SubprocessRFT_Detail = new List<Production_SubprocessRFT>(),
                    TestResult_Detail = new Production_TestResult(),
                    InlineRFT_Detail = new List<Production_InlineRFT>(),
                    EndlineWFT_Detail = new List<Production_EndlineWFT>(),
                    MDPassRate_Detail = new List<Production_MDPassRate>()
                },
                FinalInspection = new DatabaseObject.ResultModel.EtoEFlowChart.FinalInspection(),
            };
            return View(model);
        }
        [HttpPost]
        public ActionResult Index(EtoEFlowChart_Request request)
        {
            EtoEFlowChart_ViewModel model = new EtoEFlowChart_ViewModel()
            {
                Request = new EtoEFlowChart_Request(),
                Warehouse = new Warehouse()
                {
                    PhysicalInspection_Detail = new List<Warehouse_PhysicalInspection>(),
                },
                Development = new Development()
                {
                    SampleRFT_Detail = new List<Development_SampleRFT>(),
                    TestResult_Detail = new List<Development_TestResult>(),
                    RRLR_Detail = new List<Development_RRLR>(),
                    FD_Detail = new List<Development_FD>(),
                    CFTComments = new List<Development_SampleRFT_CFTComments>()
                },
                Production = new Production()
                {
                    SubprocessRFT_Detail = new List<Production_SubprocessRFT>(),
                    TestResult_Detail = new Production_TestResult(),
                    InlineRFT_Detail = new List<Production_InlineRFT>(),
                    EndlineWFT_Detail = new List<Production_EndlineWFT>(),
                    MDPassRate_Detail = new List<Production_MDPassRate>()
                },
                FinalInspection = new DatabaseObject.ResultModel.EtoEFlowChart.FinalInspection(),
            };

            if (string.IsNullOrEmpty(request.BrandID) || string.IsNullOrEmpty(request.SeasonID) || string.IsNullOrEmpty(request.StyleID))
            {
                model.ErrorMessage = $@"msg.WithInfo(""BrandID, Season, Style can't be empty."");";
                return View(model);
            }

            model = _Service.GetDate(request);
            model.Request = request;

            if (!model.ExecuteResult)
            {
                string ErrorMessage = $@"msg.WithInfo(`{model.ErrorMessage}`);";
                model = new EtoEFlowChart_ViewModel()
                {
                    Request = new EtoEFlowChart_Request(),
                    Warehouse = new Warehouse()
                    {
                        PhysicalInspection_Detail = new List<Warehouse_PhysicalInspection>(),
                    },
                    Development = new Development()
                    {
                        SampleRFT_Detail = new List<Development_SampleRFT>(),
                        TestResult_Detail = new List<Development_TestResult>(),
                        RRLR_Detail = new List<Development_RRLR>(),
                        FD_Detail = new List<Development_FD>(),
                        CFTComments = new List<Development_SampleRFT_CFTComments>()
                    },
                    Production = new Production()
                    {
                        SubprocessRFT_Detail = new List<Production_SubprocessRFT>(),
                        TestResult_Detail = new Production_TestResult(),
                        InlineRFT_Detail = new List<Production_InlineRFT>(),
                        EndlineWFT_Detail = new List<Production_EndlineWFT>(),
                        MDPassRate_Detail = new List<Production_MDPassRate>()
                    },
                    FinalInspection = new DatabaseObject.ResultModel.EtoEFlowChart.FinalInspection(),
                };
                model.ErrorMessage = ErrorMessage;
            }
            return View(model);
        }
    }
}