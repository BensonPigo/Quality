
﻿
using BusinessLogicLayer.Service.FinalInspection;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using FactoryDashBoardWeb.Helper;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.FinalInspection.Controllers
{
    public class QueryController : Controller
    {
        private QueryService Service = new QueryService();

        // GET: FinalInspection/Query
        public ActionResult Index()
        {

            List<string> inspectionlist = new List<string>() {
                "", "Pass","Fail","On-going"
            };

            List<SelectListItem> inspectionResultList = new SetListItem().ItemListBinding(inspectionlist);
            ViewBag.inspectionResultList = inspectionResultList;

            List<DatabaseObject.ViewModel.FinalInspection.QueryFinalInspection> list = new List<DatabaseObject.ViewModel.FinalInspection.QueryFinalInspection>();
            DatabaseObject.ViewModel.FinalInspection.QueryFinalInspection temp = new DatabaseObject.ViewModel.FinalInspection.QueryFinalInspection();
            DatabaseObject.ViewModel.FinalInspection.QueryFinalInspection temp2 = new DatabaseObject.ViewModel.FinalInspection.QueryFinalInspection();
            for (int i = 0; i < 25; i++)
            {
                temp.SP = "ID" + i;
                temp.POID = "POID1";
                temp.SPQty = "123";
                temp.StyleID = "styleID1";
                temp.Season = "SeasonID1";
                temp.BrandID = "BrandID1";
                temp.InspectionTimes = "1";
                temp.InspectionStage = "Inline";
                temp.InspectionResult = "On-Going";
                list.Add(temp);
            }

            temp2.SP = "ID2";
            temp2.POID = "POID2";
            temp2.SPQty = "321";
            temp2.StyleID = "styleID1";
            temp2.Season = "SeasonID1";
            temp2.BrandID = "BrandID1";
            temp2.InspectionTimes = "2";
            temp2.InspectionStage = "2";
            temp2.InspectionResult = "Fail";
            list.Add(temp2);


            return View(list);
        }

        public ActionResult Detail()
        {
            return View();
        }

        public ActionResult DownLoad()
        {

            return View();
        }
    }
}