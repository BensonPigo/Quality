using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ResultModel.FinalInspection;
using DatabaseObject.ViewModel.FinalInspection;
using FactoryDashBoardWeb.Helper;
using Newtonsoft.Json;
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

        #region 查詢SP#
        public ActionResult Index(PoSelect Req)
        {
            this.CheckSession();

            FinalInspection_Request finalInspection_Request = new FinalInspection_Request() {
                SP = Req.SP,
                POID = Req.POID,
                StyleID = Req.StyleID,
                FactoryID = this.FactoryID
            };
            FinalInspectionService finalInspectionService = new FinalInspectionService();

            PoSelect model = new PoSelect() { DataList = new List<PoSelect_Result>() };

            List<PoSelect_Result> list = new List<PoSelect_Result>();
            string ErrorMessage = string.Empty;

            try
            {
                if (!string.IsNullOrEmpty(Req.SP) ||
                    !string.IsNullOrEmpty(Req.POID) ||
                    !string.IsNullOrEmpty(Req.StyleID))
                {
                    list = finalInspectionService.GetOrderForInspection_ByModel(finalInspection_Request).ToList();
                }
                model.DataList = list;
            }
            catch (Exception ex)
            {
                model.ErrorMessage = $@"
msg.WithInfo('{ex.Message}');
";
            }

            return View(model);
        }

        [HttpPost]
        public ActionResult Setting(List<PoSelect_Result> model, string finalInspectionID)
        {
            ViewBag.FactoryID = this.FactoryID;
            //string finalInspectionID = "";
            DatabaseObject.ViewModel.FinalInspection.Setting setting = new DatabaseObject.ViewModel.FinalInspection.Setting();
            FinalInspectionSettingService finalInspectionSettingService = new FinalInspectionSettingService();

            // 判斷是由Index頁進來，還是其他地方
            if (model != null)
            {
                var selected = model.Where(o => o.Selected).ToList();

                if (selected.Any())
                {
                    string poID = selected[0].POID;
                    List<string> listOrderID = selected.Select(s => s.ID).ToList();

                    setting = finalInspectionSettingService.GetSettingForInspection(poID, listOrderID, this.FactoryID);

                }
            }
            else
            {
                setting = finalInspectionSettingService.GetSettingForInspection(finalInspectionID);
            }

            ViewBag.InspectionStageList = new List<SelectListItem>()
            {
                new SelectListItem(){Text="Inline",Value="Inline"},
                new SelectListItem(){Text="Stagger",Value="Stagger"},
                new SelectListItem(){Text="Final",Value="Final"},
                new SelectListItem(){Text="3rd Party",Value="3rd Party"},
            };

            ViewBag.AQLPlanList = new List<SelectListItem>()
            {
                new SelectListItem(){Text="",Value=""},
                new SelectListItem(){Text="1.0 Level I",Value="1.0 Level I"},
                new SelectListItem(){Text="1.0 Level II",Value="1.0 Level II"},
                new SelectListItem(){Text="1.5 Level I",Value="1.5 Level I"},
                new SelectListItem(){Text="2.5 Level I",Value="2.5 Level I"},
                new SelectListItem(){Text="100% Inspection",Value="100% Inspection"},
            };

            //setting.SelectCarton = new List<DatabaseObject.ViewModel.FinalInspection.SelectCarton>();
            //for (int i = 1; i < 15; i++)
            //{
            //    setting = finalInspectionSettingService.GetSettingForInspection(finalInspectionID);
            //}

            TempData["Setting"] = setting;
            return View(setting);
        }

        #endregion

        #region Setting頁


        public ActionResult Setting(string finalInspectionID)
        {
            ViewBag.FactoryID = this.FactoryID;
            //string finalInspectionID = "";
            DatabaseObject.ViewModel.FinalInspection.Setting setting = new DatabaseObject.ViewModel.FinalInspection.Setting();
            FinalInspectionSettingService finalInspectionSettingService = new FinalInspectionSettingService();

            setting = finalInspectionSettingService.GetSettingForInspection(finalInspectionID);

            ViewBag.InspectionStageList = new List<SelectListItem>()
            {
                new SelectListItem(){Text="Inline",Value="Inline"},
                new SelectListItem(){Text="Stagger",Value="Stagger"},
                new SelectListItem(){Text="Final",Value="Final"},
                new SelectListItem(){Text="3rd Party",Value="3rd Party"},
            };

            ViewBag.AQLPlanList = new List<SelectListItem>()
            {
                new SelectListItem(){Text="",Value=""},
                new SelectListItem(){Text="1.0 Level I",Value="1.0 Level I"},
                new SelectListItem(){Text="1.0 Level II",Value="1.0 Level II"},
                new SelectListItem(){Text="1.5 Level I",Value="1.5 Level I"},
                new SelectListItem(){Text="2.5 Level I",Value="2.5 Level I"},
                new SelectListItem(){Text="100% Inspection",Value="100% Inspection"},
            };

            //setting.SelectCarton = new List<DatabaseObject.ViewModel.FinalInspection.SelectCarton>();
            //for (int i = 1; i < 15; i++)
            //{
            //    setting = finalInspectionSettingService.GetSettingForInspection(finalInspectionID);
            //}

            TempData["Setting"] = setting;
            ViewData["TotalAvailableQty"] = setting.SelectedPO.Sum(o => o.AvailableQty);
            ViewData["RejectQty"] = setting.AcceptQty + 1;

            return View(setting);
        }

        [HttpPost]
        public ActionResult AQL_AJAX(string AQLPlan)
        {
            Setting setting = new Setting();
            if (TempData["Setting"] != null)
            {
                setting = (Setting)TempData["Setting"];
            }
            TempData["Setting"] = setting;
            var SamplePlanQty = 0;
            var AcceptedQty = 0;
            var RejectQty = 0;
            AcceptableQualityLevels tmp = new AcceptableQualityLevels();
            switch (AQLPlan)
            {
                case "1.0 Level I":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "1" && o.LotSize_Start <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    break;
                case "1.0 Level II":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "2" && o.LotSize_Start <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    break;
                case "1.5 Level I":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == Convert.ToDecimal(1.5) && o.InspectionLevels == "1" && o.LotSize_Start <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    break;
                case "2.5 Level I":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == Convert.ToDecimal(2.5) && o.InspectionLevels == "1" && o.LotSize_Start <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    break;
                default:
                    break;
            }
            RejectQty = AcceptedQty + 1;

            var jsonObject = new List<object>();
            jsonObject.Add(JsonConvert.SerializeObject(SamplePlanQty));
            jsonObject.Add(JsonConvert.SerializeObject(AcceptedQty));
            jsonObject.Add(JsonConvert.SerializeObject(RejectQty));

            return Json(jsonObject);
        }

        /// <summary>
        /// 按下Next
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult GoGeneral(Setting setting)
        {
            FinalInspectionSettingService sevice = new FinalInspectionSettingService();

            string finalInspectionID = string.Empty;

            // Setting存檔，並取得 finalInspectionID
            BaseResult result = sevice.UpdateFinalInspection(setting, this.UserID, this.FactoryID, this.MDivisionID, out finalInspectionID);

            // 錯誤回到Setting頁
            if (!result)
            {
                setting.ErrorMessage = result.ErrorMessage;
                

                ViewBag.InspectionStageList = new List<SelectListItem>()
                {
                    new SelectListItem(){Text="Inline",Value="Inline"},
                    new SelectListItem(){Text="Stagger",Value="Stagger"},
                    new SelectListItem(){Text="Final",Value="Final"},
                    new SelectListItem(){Text="3rd Party",Value="3rd Party"},
                };

                ViewBag.AQLPlanList = new List<SelectListItem>()
                {
                    new SelectListItem(){Text="",Value=""},
                    new SelectListItem(){Text="1.0 Level I",Value="1.0 Level I"},
                    new SelectListItem(){Text="1.0 Level II",Value="1.0 Level II"},
                    new SelectListItem(){Text="1.5 Level I",Value="1.5 Level I"},
                    new SelectListItem(){Text="2.5 Level I",Value="2.5 Level I"},
                    new SelectListItem(){Text="100% Inspection",Value="100% Inspection"},
                };

                return View("Setting", setting);
            }

            return RedirectToAction("General", new { FinalInspectionID= finalInspectionID });
        }

        #endregion

        #region General頁面

        public ActionResult General(string FinalInspectionID)
        {
            FinalInspectionService sevice = new FinalInspectionService();
            DatabaseObject.ManufacturingExecutionDB.FinalInspection model = sevice.GetFinalInspection(FinalInspectionID);

            return View(model);
        }

        /// <summary>
        /// 按下Next or Back
        /// </summary>
        /// <param name="model"></param>
        /// <param name="goPage"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult GoCheckList(DatabaseObject.ManufacturingExecutionDB.FinalInspection model, string goPage)
        {
            string FinalInspectionID = model.ID;

            if (goPage == "Back")
            {
                return RedirectToAction("Setting", new { finalInspectionID = FinalInspectionID });
            }
            else if (goPage == "Next")
            {

                FinalInspectionService sevice = new FinalInspectionService();
                model.InspectionStep = "Insp-General";
                sevice.UpdateFinalInspectionByStep(model, "Insp-General", this.UserID);

                return RedirectToAction("CheckList", new { FinalInspectionID = FinalInspectionID });
            }

            return View(model);
        }
        #endregion

        #region CheckList頁面
        public ActionResult CheckList(string FinalInspectionID)
        {
            FinalInspectionService sevice = new FinalInspectionService();

            DatabaseObject.ManufacturingExecutionDB.FinalInspection model = sevice.GetFinalInspection(FinalInspectionID);

            return View(model);
        }

        [HttpPost]
        public ActionResult CheckList(DatabaseObject.ManufacturingExecutionDB.FinalInspection finalinspection, string goPage)
        {
            if (goPage == "Back")
            {
                return RedirectToAction("General", new { FinalInspectionID = finalinspection.ID });
            }
            else if (goPage == "Next")
            {
                FinalInspectionService sevice = new FinalInspectionService();
                finalinspection.InspectionStep = "Insp-CheckList";
                sevice.UpdateFinalInspectionByStep(finalinspection, "Insp-CheckList", this.UserID);

                return RedirectToAction("AddDefect", new { FinalInspectionID = finalinspection.ID });
            }
            return View();
        }
        #endregion

        #region AddDefect頁面
        #endregion
        public ActionResult AddDefect(string FinalInspectionID)
        {
            FinalInspectionService sevice = new FinalInspectionService();
            FinalInspectionAddDefectService Addsevice = new FinalInspectionAddDefectService();

            DatabaseObject.ManufacturingExecutionDB.FinalInspection model = sevice.GetFinalInspection(FinalInspectionID);

            DatabaseObject.ViewModel.FinalInspection.AddDefect addDefct = new DatabaseObject.ViewModel.FinalInspection.AddDefect();
            addDefct = Addsevice.GetDefectForInspection(FinalInspectionID);

            addDefct.FinalInspectionID = FinalInspectionID;
            addDefct.RejectQty = model.RejectQty;
            addDefct.SampleSize = model.SampleSize;


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
        public ActionResult AddDefect(DatabaseObject.ViewModel.FinalInspection.AddDefect addDefct, string goPage)
        {
            if (goPage == "Back")
            {
                return RedirectToAction("CheckList");
            }
            else if (goPage == "Next")
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
            //for (int i = 1; i < 10; i++)
            //{
            //    DatabaseObject.ViewModel.FinalInspection.CartonItem cartonItem = new DatabaseObject.ViewModel.FinalInspection.CartonItem();
            //    cartonItem.FinalInspection_OrderCartonUkey = i + 1;
            //    cartonItem.OrderID = "A001";
            //    cartonItem.CTNNo = i.ToString();
            //    cartonItem.PackinglistID = i.ToString();

            //    moisture.ListCartonItem.Add(cartonItem);
            //}

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
        [HttpPost]
        public ActionResult Moisture(DatabaseObject.ViewModel.FinalInspection.MoistureResult moistureResult, string goPage)
        {
            if (goPage == "Back")
            {
                return RedirectToAction("BeautifulProductAudit");
            }
            else if (goPage == "Next")
            {
                return RedirectToAction("Measurement");
            }
            return View();
        }

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

            Measurement.SizeUnit = "INCH";

            DatabaseObject.ViewModel.FinalInspection.MeasurementItem MeasurementItem = new DatabaseObject.ViewModel.FinalInspection.MeasurementItem();
            MeasurementItem.Description = "AAA";
            MeasurementItem.Tol1 = "444 15/26";
            MeasurementItem.Tol2 = "999 13/26";
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

            return View(Measurement);
        }

        [HttpPost]
        public ActionResult Measurement(DatabaseObject.ViewModel.FinalInspection.Measurement Measurement, string goPage)
        {
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
        public ActionResult Others(DatabaseObject.ViewModel.FinalInspection.Others others, string goPage)
        {
            if (goPage == "Back")
            {
                return RedirectToAction("Measurement");
            }
            else if (goPage == "Submit")
            {
                return RedirectToAction("Others");
            }
            return View(); ;
        }


    }
}