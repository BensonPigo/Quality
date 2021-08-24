using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service;
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
        // 測試前端假資料格式
        public class ListEndlineMoistureClass
        {
            public string Instrument;
            public string Fabrication;
            public decimal Standard;
        }

        public class ActionClass
        {
            public string name;
        }

        // Setting
        public ActionResult Index(string txtSP, string txtPO, string txtStyle)
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
            ViewBag.FactoryID = this.FactoryID;

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
            for (int i = 1; i < 20; i++)
            {
                DatabaseObject.ViewModel.FinalInspection.SelectedPO item = new DatabaseObject.ViewModel.FinalInspection.SelectedPO();
                item.OrderID = "A00" + i.ToString();
                item.POID = i.ToString();
                item.StyleID = "2";
                item.SeasonID = "3";
                item.BrandID = "4";
                item.Qty = i * 2;
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
            if (page == "Back")
            {
                return RedirectToAction("Setting");
            }
            else if (page == "Next")
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

        public List<SelectListItem> ItemListBinding(Dictionary<string, string> Options)
        {
            List<SelectListItem> result_itemList = new List<SelectListItem>();
            foreach (var item in Options)
            {
                SelectListItem i = new SelectListItem()
                {
                    Text = item.Value,
                    Value = item.Key,
                };
                result_itemList.Add(i);
            }

            return result_itemList;
        }

        public ActionResult Moisture()
        {
            DatabaseObject.ViewModel.FinalInspection.Moisture moisture = new DatabaseObject.ViewModel.FinalInspection.Moisture();
            moisture.FinalInspectionID = "ESPCH21080001";
            moisture.FinalInspection_CTNMoisureStandard = 7.5m;
            moisture.ListArticle = new List<string>();
            for (int i= 10;i<60;i=i+10)
            {
                moisture.ListArticle.Add("00" + i.ToString());
            }
            moisture.ListCartonItem = new List<DatabaseObject.ViewModel.FinalInspection.CartonItem>();
            for (int i = 1; i < 10; i++)
            {
                DatabaseObject.ViewModel.FinalInspection.CartonItem cartonItem = new DatabaseObject.ViewModel.FinalInspection.CartonItem();
                cartonItem.FinalInspection_OrderCartonUkey = i;
                cartonItem.OrderID = "A001";
                cartonItem.CTNNo = i.ToString();
                cartonItem.PackinglistID = i.ToString();

                moisture.ListCartonItem.Add(cartonItem);
            }
            moisture.ListEndlineMoisture = new List<DatabaseObject.ManufacturingExecutionDB.EndlineMoisture>();

            string jsonString = @"
[
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '100% Cotton',
    'Standard': 56
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '100% Linen',
    'Standard': 67
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '100% Polyacrylic',
    'Standard': 83
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '100% Polyamide',
    'Standard': 67
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '100% Polyester',
    'Standard': 57
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '100% Viscose',
    'Standard': 59
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '30% Viscose 70% Polyester',
    'Standard': 62
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '50% Cotton 50% Polyacrylic',
    'Standard': 62
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '50% Cotton 50% Polyester',
    'Standard': 37
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '50% Viscose 50% Cotton',
    'Standard': 48
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '50% Viscose 50% Polyester',
    'Standard': 57
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '60% Cotton 40% Polyester',
    'Standard': 45
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '60% Linen 40% Cotton',
    'Standard': 47
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '70% Cotton 30% Polyamide',
    'Standard': 59
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '70% Cotton 30% Polyester',
    'Standard': 53
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '80% Cotton 20% Polyester',
    'Standard': 59
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '80% Viscose 20% Polyamide',
    'Standard': 59
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '90% Cotton 10% Elastane/Spandex',
    'Standard': 56
  },
  {
    'Instrument': 'Aqua Boy',
    'Fabrication': '90% Viscose 10% Elastane/Spandex',
    'Standard': 59
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Acetate',
    'Standard': 7.5
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Acrylic',
    'Standard': 5.5
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Cotton',
    'Standard': 8
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Leather',
    'Standard': 4.6
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Linen/Flax',
    'Standard': 9.5
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Nylon',
    'Standard': 8
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Paper/Straw',
    'Standard': 6.6
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Polyester',
    'Standard': 5.5
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'PU',
    'Standard': 4.6
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Silk',
    'Standard': 6
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Viscose',
    'Standard': 6
  },
  {
    'Instrument': 'B Machine',
    'Fabrication': 'Wool',
    'Standard': 4.5
  }
]";
            List<ListEndlineMoistureClass> objectList = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ListEndlineMoistureClass>>(jsonString);
            foreach (ListEndlineMoistureClass item in objectList)
            {
                DatabaseObject.ManufacturingExecutionDB.EndlineMoisture endlineMoisture = new DatabaseObject.ManufacturingExecutionDB.EndlineMoisture();
                endlineMoisture.Instrument = item.Instrument;
                endlineMoisture.Fabrication = item.Fabrication;
                endlineMoisture.Standard = item.Standard;

                moisture.ListEndlineMoisture.Add(endlineMoisture);
            }

            moisture.ActionSelectListItem = new List<SelectListItem>();
            jsonString = @"[{'name':''},{'name':'Change carton'},{'name':'Drying garment+change carton'},{'name':'Open carton with drying garment'},{'name':'Others'}]";
            List<ActionClass> actionList = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ActionClass>>(jsonString);
            foreach (ActionClass item in actionList)
            {
                SelectListItem i = new SelectListItem { Text = item.name, Value = item.name };
                moisture.ActionSelectListItem.Add(i);
            }

            ViewBag.ListArticle = new SetListItem().ItemListBinding(moisture.ListArticle);
            ViewBag.ListCartonItem = moisture.ListCartonItem;
            ViewBag.ListEndlineMoisture = moisture.ListEndlineMoisture;
            ViewBag.ActionSelectListItem = moisture.ActionSelectListItem;

            ViewBag.FinalInspectionID = moisture.FinalInspectionID;
            ViewBag.FinalInspection_CTNMoisureStandard = moisture.FinalInspection_CTNMoisureStandard;

            DatabaseObject.ViewModel.FinalInspection.MoistureResult moistureResult = new DatabaseObject.ViewModel.FinalInspection.MoistureResult();
            return View(moistureResult);
        }

        [HttpPost]
        public ActionResult GetViewMoistureResult(string finalInspectionID)
        {
            IFinalInspectionMoistureService finalInspectionMoistureService = new FinalInspectionMoistureService();
            List<DatabaseObject.ViewModel.FinalInspection.ViewMoistureResult> viewMoistureResultsList = finalInspectionMoistureService.GetViewMoistureResult(finalInspectionID);

            return Json(viewMoistureResultsList);
        }

        // GET: SampleRFT/Measurement
        public ActionResult Measurement()
        {

            List<string> listArticle = new List<string>() {
                 "0050","0049"
            };

            List<SelectListItem> articleList = new SetListItem().ItemListBinding(listArticle);
            ViewBag.ListArticle = articleList;

            List<string> listSize = new List<string>() {
                 "S","M","L","XL"
            };

            List<SelectListItem> sizeList = new SetListItem().ItemListBinding(listSize);
            ViewBag.ListSize = sizeList;


            List<string> listProductType = new List<string>() {
                 "Bottom","Top"
            };

            List<SelectListItem> productTypeList = new SetListItem().ItemListBinding(listProductType);
            ViewBag.ListProductType = productTypeList;

            DatabaseObject.ViewModel.FinalInspection.Measurement Measurement = new DatabaseObject.ViewModel.FinalInspection.Measurement();

            Measurement.SizeUnit = "CM";

            DatabaseObject.ViewModel.FinalInspection.MeasurementItem MeasurementItem = new DatabaseObject.ViewModel.FinalInspection.MeasurementItem();
            MeasurementItem.Description = "AAA";
            MeasurementItem.Tol1 = "11";
            MeasurementItem.Tol2 = "22";
            MeasurementItem.SizeSpec = "44";
            MeasurementItem.Size = "M";
            MeasurementItem.ResultSizeSpec = "";
            MeasurementItem.CanEdit = true;
            Measurement.ListMeasurementItem = new List<DatabaseObject.ViewModel.FinalInspection.MeasurementItem>();
            Measurement.ListMeasurementItem.Add(MeasurementItem);

            DatabaseObject.ViewModel.FinalInspection.MeasurementItem MeasurementItem2 = new DatabaseObject.ViewModel.FinalInspection.MeasurementItem();
            MeasurementItem2.Description = "BBB";
            MeasurementItem2.Tol1 = "11";
            MeasurementItem2.Tol2 = "22";
            MeasurementItem2.SizeSpec = "12.6";
            MeasurementItem2.Size = "L";
            MeasurementItem2.ResultSizeSpec = "";
            MeasurementItem2.CanEdit = false;

            Measurement.ListMeasurementItem.Add(MeasurementItem2);

            DatabaseObject.ViewModel.FinalInspection.MeasurementItem MeasurementItem3 = new DatabaseObject.ViewModel.FinalInspection.MeasurementItem();
            MeasurementItem3.Description = "CCC";
            MeasurementItem3.Tol1 = "11";
            MeasurementItem3.Tol2 = "22";
            MeasurementItem3.SizeSpec = "12";
            MeasurementItem3.Size = "L";
            MeasurementItem3.ResultSizeSpec = "";
            MeasurementItem3.CanEdit = true;

            Measurement.ListMeasurementItem.Add(MeasurementItem3);


            //TempData["Model"] = null;
            return View(Measurement);
        }

        [HttpPost]
        public ActionResult Measurement(DatabaseObject.ViewModel.FinalInspection.Measurement Measurement, string goPage)
        {
            //string someJson = "[  {    \"Code\": \"S021\",    \"Description\": \"A Chest width ( meas. 2cm below armhole )\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"103\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S03\",    \"Description\": \"B WAIST WIDTH\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"99\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S04\",    \"Description\": \"B1 WAIST MEAS. POINT FROM CHEST MEAS. POINT\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"19\",    \"2021/08/20 16:25:38\": \"1\",    \"diff1\": \"-18\"  },  {    \"Code\": \"S05\",    \"Description\": \"D HEM OPENING  (MEAS. STRAIGHT)\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"102\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S06\",    \"Description\": \"E Front zip length (zipend 0,5cm before collar edge) (tolerance +/- 1%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"26\",    \"2021/08/20 16:25:38\": \"2\",    \"diff1\": \"-24\"  },  {    \"Code\": \"S07\",    \"Description\": \"F FRONT TO BACK\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"3\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S08\",    \"Description\": \"H 1/2 ZIP - Shoulder length\",    \"Tol(+)\": \"0.8\",    \"Tol(-)\": \"0.4\",    \"50_aa\": \"14.6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S09\",    \"Description\": \"I 2 PIECE - PRESHAPE  sleeve - Sleeve length\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"67\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S10\",    \"Description\": \"J Sleeve width ( meas. 2cm below armhole )\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"39.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S11\",    \"Description\": \"K LONG + PRESHAPE SLEEVE -Ellbow width (meas. 32,0cm above sleeve opening)\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"28.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S14\",    \"Description\": \"M LONG + PRESHAPE sleeve - Sleeve opening\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"21\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S15\",    \"Description\": \"O1 1/2 ZIP - Front neck drop (HPS to c.f. neck seam)\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0\",    \"50_aa\": \"10.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S16\",    \"Description\": \"O2 1/2 ZIP - Back neck drop (HPS to c.b. neck seam) pattern meas.\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0\",    \"50_aa\": \"2\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S17\",    \"Description\": \"N 1/2 ZIP - Back neck width (HPS to HPS)\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"14.4\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S18\",    \"Description\": \"Q1 COLLAR LENGTH OUTER EDGE\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"41.6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S19\",    \"Description\": \"Q2 Collar height (center front) (tolerance +/- 10%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S20\",    \"Description\": \"Q3 COLLAR HEIGHT (CENTER BACK) (TOLERANCE +/- 10%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"4.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S21\",    \"Description\": \"S MINIMUM NECK OPENING STRETCHED\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"63\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S22\",    \"Description\": \"Y1 Logo meas.: top edge of logo meas. to HPS\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"17.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S23\",    \"Description\": \"Y2 LOGO MEAS.: EDGE OF LOGO TO CENTER FRONT\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"7\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S24\",    \"Description\": \"Z 1/2 ZIP - Back length\",    \"Tol(+)\": \"1.5\",    \"Tol(-)\": \"1\",    \"50_aa\": \"74\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S25\",    \"Description\": \"DEC LABEL\",    \"Tol(+)\": \"\",    \"Tol(-)\": \"\",    \"50_aa\": \"50\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S27\",    \"Description\": \"POLYBAG SIZE (WXL)\",    \"Tol(+)\": \"\",    \"Tol(-)\": \"\",    \"50_aa\": \"30X40\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  }]";
            //   _MeasurementService = new MeasurementService();

            //Measurement_Request measurementRequest = _MeasurementService.MeasurementGetPara(request.OrderID, this.FactoryID);
            //Measurement_ResultModel measurement = _MeasurementService.MeasurementGet(measurementRequest);

            //TempData["Model"] = someJson;// measurement.JsonBody;
            //measurement.JsonBody = someJson;
            //return View(measurement);

            if (goPage == "Back")
            {
                return RedirectToAction("Moisture");
            }
            else if (goPage == "Next")
            {
                return RedirectToAction("Others");
            }

            return View();
        }

        //public ActionResult Measurement()
        //{
        //    return View();
        //}


        [HttpPost]
        public ActionResult OpenView()
        {
            List<DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem> listMeasurementViewItem = new List<DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem>();

            DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem MeasurementViewItem1 = new DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem();
            MeasurementViewItem1.Article = "0050";
            MeasurementViewItem1.Size = "L";
            MeasurementViewItem1.ProductType = "Top";

            listMeasurementViewItem.Add(MeasurementViewItem1);

            DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem MeasurementViewItem2 = new DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem();
            MeasurementViewItem2.Article = "0050";
            MeasurementViewItem2.Size = "XL";
            MeasurementViewItem2.ProductType = "Bottom";

            listMeasurementViewItem.Add(MeasurementViewItem2);

            DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem MeasurementViewItem3 = new DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem();
            MeasurementViewItem3.Article = "0049";
            MeasurementViewItem3.Size = "S";
            MeasurementViewItem3.ProductType = "Top";


            DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem MeasurementViewItem4 = new DatabaseObject.ViewModel.FinalInspection.MeasurementViewItem();
            MeasurementViewItem4.Article = "0049";
            MeasurementViewItem4.Size = "M";
            MeasurementViewItem4.ProductType = "Top";

            listMeasurementViewItem.Add(MeasurementViewItem4);




            //string someJson = "[  {    \"Code\": \"S021\",    \"Description\": \"A Chest width ( meas. 2cm below armhole )\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"103\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S03\",    \"Description\": \"B WAIST WIDTH\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"99\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S04\",    \"Description\": \"B1 WAIST MEAS. POINT FROM CHEST MEAS. POINT\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"19\",    \"2021/08/20 16:25:38\": \"1\",    \"diff1\": \"-18\"  },  {    \"Code\": \"S05\",    \"Description\": \"D HEM OPENING  (MEAS. STRAIGHT)\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"102\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S06\",    \"Description\": \"E Front zip length (zipend 0,5cm before collar edge) (tolerance +/- 1%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"26\",    \"2021/08/20 16:25:38\": \"2\",    \"diff1\": \"-24\"  },  {    \"Code\": \"S07\",    \"Description\": \"F FRONT TO BACK\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"3\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S08\",    \"Description\": \"H 1/2 ZIP - Shoulder length\",    \"Tol(+)\": \"0.8\",    \"Tol(-)\": \"0.4\",    \"50_aa\": \"14.6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S09\",    \"Description\": \"I 2 PIECE - PRESHAPE  sleeve - Sleeve length\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"67\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S10\",    \"Description\": \"J Sleeve width ( meas. 2cm below armhole )\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"39.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S11\",    \"Description\": \"K LONG + PRESHAPE SLEEVE -Ellbow width (meas. 32,0cm above sleeve opening)\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"28.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S14\",    \"Description\": \"M LONG + PRESHAPE sleeve - Sleeve opening\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"21\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S15\",    \"Description\": \"O1 1/2 ZIP - Front neck drop (HPS to c.f. neck seam)\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0\",    \"50_aa\": \"10.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S16\",    \"Description\": \"O2 1/2 ZIP - Back neck drop (HPS to c.b. neck seam) pattern meas.\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0\",    \"50_aa\": \"2\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S17\",    \"Description\": \"N 1/2 ZIP - Back neck width (HPS to HPS)\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"14.4\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S18\",    \"Description\": \"Q1 COLLAR LENGTH OUTER EDGE\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"41.6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S19\",    \"Description\": \"Q2 Collar height (center front) (tolerance +/- 10%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S20\",    \"Description\": \"Q3 COLLAR HEIGHT (CENTER BACK) (TOLERANCE +/- 10%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"4.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S21\",    \"Description\": \"S MINIMUM NECK OPENING STRETCHED\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"63\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S22\",    \"Description\": \"Y1 Logo meas.: top edge of logo meas. to HPS\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"17.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S23\",    \"Description\": \"Y2 LOGO MEAS.: EDGE OF LOGO TO CENTER FRONT\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"7\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S24\",    \"Description\": \"Z 1/2 ZIP - Back length\",    \"Tol(+)\": \"1.5\",    \"Tol(-)\": \"1\",    \"50_aa\": \"74\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S25\",    \"Description\": \"DEC LABEL\",    \"Tol(+)\": \"\",    \"Tol(-)\": \"\",    \"50_aa\": \"50\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S27\",    \"Description\": \"POLYBAG SIZE (WXL)\",    \"Tol(+)\": \"\",    \"Tol(-)\": \"\",    \"50_aa\": \"30X40\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  }]";

            return Json(listMeasurementViewItem);
        }


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
        public ActionResult Others(DatabaseObject.ViewModel.FinalInspection.Others others, string page)
        {
            if (page == "Back")
            {
                return RedirectToAction("Measurement");
            }
            else if (page == "Submit")
            {
                return RedirectToAction("Others");
            }
            return View(); ;
        }


    }
}