using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using FactoryDashBoardWeb.Helper;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.FinalInspection.Controllers
{
    public class InspectionController : BaseController
    {
        // Setting
        public ActionResult Index(string txtSP,string txtPO,string txtStyle)
        {

            List<DatabaseObject.ProductionDB.Orders> list = new List<DatabaseObject.ProductionDB.Orders>();
            DatabaseObject.ProductionDB.Orders temp = new DatabaseObject.ProductionDB.Orders();
            DatabaseObject.ProductionDB.Orders temp2 = new DatabaseObject.ProductionDB.Orders();
            for (int i = 0; i < 25; i++)
            {
                temp.ID = "ID" + i;
                temp.POID = "POID1";
                temp.Qty = 123;
                temp.StyleID = "styleID1";
                temp.SeasonID = "SeasonID1";
                temp.BrandID = "BrandID1";
                list.Add(temp);
            }


            temp2.ID = "ID2";
            temp2.POID = "POID2";
            temp2.Qty = 234;
            temp2.StyleID = "styleID2";
            temp2.SeasonID = "SeasonID2";
            temp2.BrandID = "BrandID2";
            list.Add(temp2);


            return View(list);
        }

        [HttpPost]
        public ActionResult GoSetting(List<DatabaseObject.ProductionDB.Orders> models)
        {

            return RedirectToAction("Setting");
        }

        public ActionResult Setting()
        {
            List<string> list = new List<string>() {
                 "Inline","Stagger","Final","3rd Party"
            };
            List<SelectListItem> inspectionStageList = new SetListItem().ItemListBinding(list);
            ViewBag.InspectionStageList = inspectionStageList;

            list = new List<string>() {
                string.Empty, "1.0 Level I", "1.0 Level II", "1.5 Level I", "2.5 Level I", "100% Inspection"
            };
            List<SelectListItem> aqlPlanList = new SetListItem().ItemListBinding(list);
            ViewBag.AQLPlanList = aqlPlanList;

            DatabaseObject.ViewModel.FinalInspection.Setting setting = new DatabaseObject.ViewModel.FinalInspection.Setting();
            setting.InspectionStage = "Inline";
            setting.AuditDate = System.DateTime.Now;
            setting.SewingLineID = "8";
            setting.InspectionTimes = "122";
            setting.SampleSize = 8;
            setting.AcceptQty = 122;

            setting.SelectedPO = new List<DatabaseObject.ViewModel.FinalInspection.SelectedPO>();
            for (int i =1;i<3;i++)
            {
                DatabaseObject.ViewModel.FinalInspection.SelectedPO item = new DatabaseObject.ViewModel.FinalInspection.SelectedPO();
                item.OrderID = "A00" + i.ToString();
                item.POID = i.ToString(); 
                item.StyleID = "2";
                item.SeasonID = "3";
                item.BrandID = "4";
                item.Qty = i*2;
                item.Cartons = i.ToString();
                item.AvailableQty = i;
                setting.SelectedPO.Add(item);
            }

            setting.SelectCarton = new List<DatabaseObject.ViewModel.FinalInspection.SelectCarton>();
            for (int i = 1; i < 3; i++)
            {
                DatabaseObject.ViewModel.FinalInspection.SelectCarton item = new DatabaseObject.ViewModel.FinalInspection.SelectCarton();
                item.Selected = true;
                item.OrderID = i.ToString();
                item.PackingListID = "1";
                item.CTNNo = i.ToString();
                setting.SelectCarton.Add(item);
            }

            return View(setting);
        }

        [HttpPost]
        public ActionResult Setting(DatabaseObject.ViewModel.FinalInspection.Setting setting)
        {
            return RedirectToAction("General");
        }
        public ActionResult General()
        {
            return View();
        }

        [HttpPost]
        public ActionResult General(List<DatabaseObject.ProductionDB.Orders> model, string page)
        {
            if (page=="Back")
            {
                return RedirectToAction("Setting");
            }
            else if(page == "Next")
            {
                return RedirectToAction("CheckList");
            }

            return View();
        }

        //[HttpPost]
        //public ActionResult General(List<DatabaseObject.ProductionDB.Orders>  model)
        //{
        //    return View();
        //}


        public ActionResult CheckList()
        {
            DatabaseObject.ManufacturingExecutionDB.FinalInspection finalinspection = new DatabaseObject.ManufacturingExecutionDB.FinalInspection();

            finalinspection.ID = "A001";
            finalinspection.POID = "1";
            finalinspection.CheckHandfeel = true;
            return View(finalinspection);
        }

        [HttpPost]
        public ActionResult CheckList(DatabaseObject.ManufacturingExecutionDB.FinalInspection finalinspection, string goPage)
        {
            return View();
        }

        public ActionResult AddDefect()
        {
            DatabaseObject.ViewModel.FinalInspection.AddDefect addDefct = new DatabaseObject.ViewModel.FinalInspection.AddDefect();
            addDefct.RejectQty = 8;
            addDefct.ListFinalInspectionDefectItem = new List<DatabaseObject.ViewModel.FinalInspection.FinalInspectionDefectItem>();
            DatabaseObject.ViewModel.FinalInspection.FinalInspectionDefectItem data = new DatabaseObject.ViewModel.FinalInspection.FinalInspectionDefectItem();
            data.DefectTypeDesc = "Accessories";
            data.DefectCodeDesc = "Accessories broken/damage";
            data.Qty = 2;
            data.Ukey = 0;
            addDefct.ListFinalInspectionDefectItem.Add(data);

            DatabaseObject.ViewModel.FinalInspection.FinalInspectionDefectItem data2 = new DatabaseObject.ViewModel.FinalInspection.FinalInspectionDefectItem();
            data2.DefectTypeDesc = "Accessories";
            data2.DefectCodeDesc = "Accessories missing/uncompleted";
            data2.Qty = 0;
            data.Ukey = 1;
            addDefct.ListFinalInspectionDefectItem.Add(data2);


            DatabaseObject.ViewModel.FinalInspection.FinalInspectionDefectItem data3 = new DatabaseObject.ViewModel.FinalInspection.FinalInspectionDefectItem();
            data3.DefectTypeDesc = "Cleanliness";
            data3.DefectCodeDesc = "Cleanliness ABC";
            data3.Qty = 0;
            data.Ukey = 2;
            addDefct.ListFinalInspectionDefectItem.Add(data3);



            return View(addDefct);
        }

        [HttpPost]
        public ActionResult AddDefect(DatabaseObject.ViewModel.FinalInspection.AddDefect addDefct, string page)
        {
            if (page == "Back")
            {
                return RedirectToAction("CheckList");
            }
            else if (page == "Next")
            {
                return RedirectToAction("BeautifulProductAudit");
            }
            return View();

        }

        public ActionResult BeautifulProductAudit()
        {
            DatabaseObject.ViewModel.FinalInspection.BeautifulProductAudit beautifulProductAudit = new DatabaseObject.ViewModel.FinalInspection.BeautifulProductAudit();
            beautifulProductAudit.FinalInspectionID = "1";
            beautifulProductAudit.BAQty = 10;
            beautifulProductAudit.SampleSize = 2;
            beautifulProductAudit.ListBACriteria = new List<DatabaseObject.ViewModel.FinalInspection.BACriteriaItem>();
            DatabaseObject.ViewModel.FinalInspection.BACriteriaItem test1 = new DatabaseObject.ViewModel.FinalInspection.BACriteriaItem();
            test1.Ukey = 1;
            test1.BACriteria = "C1";
            test1.BACriteriaDesc = "Delights consumers";
            beautifulProductAudit.ListBACriteria.Add(test1);

            DatabaseObject.ViewModel.FinalInspection.BACriteriaItem test2 = new DatabaseObject.ViewModel.FinalInspection.BACriteriaItem();
            test2.Ukey = 2;
            test2.BACriteria = "C9";
            test2.BACriteriaDesc = "Well finished and presented";
            test2.Qty = 3;
            beautifulProductAudit.ListBACriteria.Add(test2);

            return View(beautifulProductAudit);
        }

        [HttpPost]
        public ActionResult BeautifulProductAudit(DatabaseObject.ViewModel.FinalInspection.BeautifulProductAudit beautifulProductAudit, string goPage)
        {

            if (goPage == "Back")
            {
                return RedirectToAction("AddDefect");
            }
            else if (goPage == "Next")
            {
                return RedirectToAction("Moisture");
            }
            return View();

        }

        public ActionResult Moisture()
        {
            DatabaseObject.ViewModel.FinalInspection.Moisture moisture = new DatabaseObject.ViewModel.FinalInspection.Moisture();
            return View();
        }

        private IMeasurementService _MeasurementService;
        public void MeasurementController()
        {
            _MeasurementService = new MeasurementService();
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.Measurement,,";
        }


        // GET: SampleRFT/Measurement
        public ActionResult Measurement()
        {
            //_MeasurementService = new MeasurementService();
            //this.SelectedMenu = "Sample RFT";
            //ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.Measurement,,";

            List<string> listArticle = new List<string>() {
                 "Inline","Stagger","Final","3rd Party"
            };

            List<SelectListItem> articleList = new SetListItem().ItemListBinding(listArticle);
            ViewBag.ListArticle = articleList;

            List<string> listSize = new List<string>() {
                 "Inlinea","Stagger","Final","3rd Party"
            };

            List<SelectListItem> sizeList = new SetListItem().ItemListBinding(listSize);
            ViewBag.ListSize = sizeList;


            List<string> listProductType = new List<string>() {
                 "Inlineb","Stagger","Final","3rd Party"
            };

            List<SelectListItem> productTypeList = new SetListItem().ItemListBinding(listProductType);
            ViewBag.ListProductType = productTypeList;

            DatabaseObject.ViewModel.FinalInspection.Measurement Measurement = new DatabaseObject.ViewModel.FinalInspection.Measurement();
            DatabaseObject.ViewModel.FinalInspection.MeasurementItem MeasurementItem = new DatabaseObject.ViewModel.FinalInspection.MeasurementItem();
            MeasurementItem.Description = "AAA";
            MeasurementItem.Tol1 = "11";
            MeasurementItem.Tol2 = "22";
            MeasurementItem.SizeSpec = "dd";
            MeasurementItem.ResultSizeSpec = "";

            Measurement.ListMeasurementItem = new List<DatabaseObject.ViewModel.FinalInspection.MeasurementItem>();
            Measurement.ListMeasurementItem.Add(MeasurementItem);

            DatabaseObject.ViewModel.FinalInspection.MeasurementItem MeasurementItem2 = new DatabaseObject.ViewModel.FinalInspection.MeasurementItem();
            MeasurementItem2.Description = "AAA";
            MeasurementItem2.Tol1 = "11";
            MeasurementItem2.Tol2 = "22";
            MeasurementItem2.SizeSpec = "dd";
            MeasurementItem2.ResultSizeSpec = "";

            Measurement.ListMeasurementItem.Add(MeasurementItem2);


            //TempData["Model"] = null;
            return View(Measurement);
        }

        [HttpPost]
        public ActionResult Measurement(DatabaseObject.ViewModel.FinalInspection.Measurement Measurement)
        {
            //string someJson = "[  {    \"Code\": \"S021\",    \"Description\": \"A Chest width ( meas. 2cm below armhole )\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"103\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S03\",    \"Description\": \"B WAIST WIDTH\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"99\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S04\",    \"Description\": \"B1 WAIST MEAS. POINT FROM CHEST MEAS. POINT\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"19\",    \"2021/08/20 16:25:38\": \"1\",    \"diff1\": \"-18\"  },  {    \"Code\": \"S05\",    \"Description\": \"D HEM OPENING  (MEAS. STRAIGHT)\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"102\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S06\",    \"Description\": \"E Front zip length (zipend 0,5cm before collar edge) (tolerance +/- 1%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"26\",    \"2021/08/20 16:25:38\": \"2\",    \"diff1\": \"-24\"  },  {    \"Code\": \"S07\",    \"Description\": \"F FRONT TO BACK\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"3\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S08\",    \"Description\": \"H 1/2 ZIP - Shoulder length\",    \"Tol(+)\": \"0.8\",    \"Tol(-)\": \"0.4\",    \"50_aa\": \"14.6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S09\",    \"Description\": \"I 2 PIECE - PRESHAPE  sleeve - Sleeve length\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"67\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S10\",    \"Description\": \"J Sleeve width ( meas. 2cm below armhole )\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"39.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S11\",    \"Description\": \"K LONG + PRESHAPE SLEEVE -Ellbow width (meas. 32,0cm above sleeve opening)\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"28.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S14\",    \"Description\": \"M LONG + PRESHAPE sleeve - Sleeve opening\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"21\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S15\",    \"Description\": \"O1 1/2 ZIP - Front neck drop (HPS to c.f. neck seam)\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0\",    \"50_aa\": \"10.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S16\",    \"Description\": \"O2 1/2 ZIP - Back neck drop (HPS to c.b. neck seam) pattern meas.\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0\",    \"50_aa\": \"2\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S17\",    \"Description\": \"N 1/2 ZIP - Back neck width (HPS to HPS)\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"14.4\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S18\",    \"Description\": \"Q1 COLLAR LENGTH OUTER EDGE\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"41.6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S19\",    \"Description\": \"Q2 Collar height (center front) (tolerance +/- 10%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S20\",    \"Description\": \"Q3 COLLAR HEIGHT (CENTER BACK) (TOLERANCE +/- 10%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"4.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S21\",    \"Description\": \"S MINIMUM NECK OPENING STRETCHED\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"63\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S22\",    \"Description\": \"Y1 Logo meas.: top edge of logo meas. to HPS\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"17.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S23\",    \"Description\": \"Y2 LOGO MEAS.: EDGE OF LOGO TO CENTER FRONT\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"7\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S24\",    \"Description\": \"Z 1/2 ZIP - Back length\",    \"Tol(+)\": \"1.5\",    \"Tol(-)\": \"1\",    \"50_aa\": \"74\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S25\",    \"Description\": \"DEC LABEL\",    \"Tol(+)\": \"\",    \"Tol(-)\": \"\",    \"50_aa\": \"50\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S27\",    \"Description\": \"POLYBAG SIZE (WXL)\",    \"Tol(+)\": \"\",    \"Tol(-)\": \"\",    \"50_aa\": \"30X40\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  }]";
            //   _MeasurementService = new MeasurementService();

            //Measurement_Request measurementRequest = _MeasurementService.MeasurementGetPara(request.OrderID, this.FactoryID);
            //Measurement_ResultModel measurement = _MeasurementService.MeasurementGet(measurementRequest);

            //TempData["Model"] = someJson;// measurement.JsonBody;
            //measurement.JsonBody = someJson;
            //return View(measurement);

            return View();
        }

        //public ActionResult Measurement()
        //{
        //    return View();
        //}

        public ActionResult Others()
        {
            DatabaseObject.ViewModel.FinalInspection.Others Others = new DatabaseObject.ViewModel.FinalInspection.Others();
            Others.CFA = "Hello";
            Others.OthersRemark = "remark";
            Others.InspectionResult = "Fail";
            Others.ShipmentStatus = "Pass";

            return View(Others);
        }


        [HttpPost]
        public ActionResult Others(DatabaseObject.ViewModel.FinalInspection.Others others,string page)
        {
            if (page == "Back")
            {
                return RedirectToAction("General");
            }
            else if (page == "Submit")
            {
                return RedirectToAction("Others");
            }
            return View();;
        }

    }
}