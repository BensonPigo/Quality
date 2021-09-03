using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class FabricCrkShrkTestController : BaseController
    {
        public FabricCrkShrkTestController()
        {
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.FabricCrkShrkTest,,";
        }

        // GET: BulkFGT/FabricCrkShrkTest
        public ActionResult Index()
        {
            DatabaseObject.ResultModel.FabricCrkShrkTest_Result result = new DatabaseObject.ResultModel.FabricCrkShrkTest_Result()
            {
                Main = new DatabaseObject.ResultModel.FabricCrkShrkTest_Main()
                {
                    POID = "123",
                    StyleID = "A",
                    BrandID = "ADIDAS",
                    SeasonID = "B",
                    CutInline = System.DateTime.Now,
                    MinSciDelivery = null,
                    TargetLeadTime = System.DateTime.Now.AddDays(-2),
                    CompletionDate = System.DateTime.Now.AddDays(-3),
                    FIRLabInspPercent = 1,
                    complete = false,
                    FirLaboratoryRemark = "Test",
                    CreateBy = "PC8000068-JESSIE BALISBIS 2019/07/11 10:43:18",
                    EditBy = "PC8000067-Test 2019/07/11 11:43:18",
                },
                Details = new List<DatabaseObject.ResultModel.FabricCrkShrkTest_Detail>()
            };

            for (int i=1;i<=5;i++)
            {
                DatabaseObject.ResultModel.FabricCrkShrkTest_Detail item = new DatabaseObject.ResultModel.FabricCrkShrkTest_Detail();
                item.Seq = i.ToString();
                item.NonCrocking = false;
                item.NonHeat = true;
                item.NonWash = false;
                item.AllResult = "Pass";
                result.Details.Add(item);
            }

            return View(result);
        }

        public ActionResult CrockingTest()
        {
            return View();
        }


        public ActionResult HeatTest()
        {
            return View();
        }


        public ActionResult WashTest()
        {
            return View();
        }
    }
}