using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web.Mvc;
using static Quality.Helper.Attribute;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ViewModel.BulkFGT;
using FactoryDashBoardWeb.Helper;
using System.Linq;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class PullingTestController : BaseController
    {
        // GET: BulkFGT/PullingTest
        public ActionResult Index()
        {
            PullingTest_ViewModel Result = new PullingTest_ViewModel()
            {
                BrandID = string.Empty,
                SeasonID = string.Empty,
                StyleID = string.Empty,
                Article = string.Empty,
                ReportNo_Source = new List<string>(),
                Detail = new PullingTest_Result()
                {
                    ReportNo = "test",

                },
            };

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(Result.ReportNo_Source);
            ViewBag.ResultList = new SetListItem().ItemListBinding("Pass,Fail".Split(',').ToList());
            return View(Result);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(PullingTest_ViewModel Req)
        {
            PullingTest_ViewModel Result = new PullingTest_ViewModel()
            {
                BrandID = Req.BrandID,
                SeasonID = Req.SeasonID,
                StyleID = Req.StyleID,
                Article = Req.Article,
                ReportNo_Source = "5,6,7".Split(',').ToList(),
                Detail = new PullingTest_Result()
                {
                    ReportNo = "PHOV190100039",
                    POID = "1234",
                    BrandID = "ADIDAS",
                    SeasonID = "19FW",
                    StyleID = "S1953WTR342",
                    Article = "FJ5381",
                    Result = "Fail",
                    TestItem = "Snaps",
                    Time = 10,
                    LastEditName = "PC8000068-JESSIE BALISBIS 2019/07/11 10:43:18"

                },
            };

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(Result.ReportNo_Source);
            ViewBag.ResultList = new SetListItem().ItemListBinding("Pass,Fail".Split(',').ToList());
            return View("Index", Result);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult EditSave(PullingTest_ViewModel Req)
        {
            PullingTest_ViewModel Result = new PullingTest_ViewModel()
            {
                BrandID = Req.BrandID,
                SeasonID = Req.SeasonID,
                StyleID = Req.StyleID,
                Article = Req.Article,
                ReportNo_Source = "1,2,3,4".Split(',').ToList(),
                Detail = new PullingTest_Result(),
            };

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(Result.ReportNo_Source);
            ViewBag.ResultList = new SetListItem().ItemListBinding("Pass,Fail".Split(',').ToList());
            return View("Index", Result);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult NewSave(PullingTest_ViewModel Req)
        {
            PullingTest_ViewModel Result = new PullingTest_ViewModel()
            {
                BrandID = Req.BrandID,
                SeasonID = Req.SeasonID,
                StyleID = Req.StyleID,
                Article = Req.Article,
                ReportNo_Source = "1,2,3,4".Split(',').ToList(),
                Detail = new PullingTest_Result(),
            };

            ViewBag.ReportNo_Source = new SetListItem().ItemListBinding(Result.ReportNo_Source);
            ViewBag.ResultList = new SetListItem().ItemListBinding("Pass,Fail".Split(',').ToList());
            return View("Index", Result);
        }
    }
}