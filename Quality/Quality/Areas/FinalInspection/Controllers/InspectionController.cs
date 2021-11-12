using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service;
using BusinessLogicLayer.Service.FinalInspection;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel.FinalInspection;
using DatabaseObject.ViewModel.FinalInspection;
using FactoryDashBoardWeb.Helper;
using Newtonsoft.Json;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web.Mvc;

namespace Quality.Areas.FinalInspection.Controllers
{
    public class InspectionController : BaseController
    {
        private string IsTest = ConfigurationManager.AppSettings["IsTest"];

        public InspectionController()
        {
            this.SelectedMenu = "Final Inspection";
            ViewBag.OnlineHelp = this.OnlineHelp + "FinalInspection.Inspection,,";
        }

        #region 查詢SP#
        public ActionResult Index(PoSelect Req)
        {
            this.CheckSession();

            FinalInspection_Request finalInspection_Request = new FinalInspection_Request()
            {
                SP = Req.SP,
                CustPONO = Req.CustPONO,
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
                    !string.IsNullOrEmpty(Req.CustPONO) ||
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
            this.CheckSession();

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
                    string CustPONO = selected[0].CustPONO;
                    List<string> listOrderID = selected.Select(s => s.ID).ToList();

                    setting = finalInspectionSettingService.GetSettingForInspection(CustPONO, listOrderID, this.FactoryID);

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

            ViewBag.Shift = new List<SelectListItem>()
            {
                new SelectListItem(){Text="Day",Value="D"},
                new SelectListItem(){Text="Night",Value="N"},
            };

            ViewBag.SewingTeam = new List<SelectListItem>();

            if (setting.SelectedSewingTeam != null)
            {
                List<SelectListItem> SewingTeam = new List<SelectListItem>();
                foreach (var item in setting.SelectedSewingTeam)
                {
                    SewingTeam.Add(new SelectListItem() { Text = item.SewingTeamID, Value = item.SewingTeamID });
                }
                ViewBag.SewingTeam = SewingTeam;
            }

            // 預設值
            setting.Team = "A";

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

            ViewBag.Shift = new List<SelectListItem>()
            {
                new SelectListItem(){Text="Day",Value="D"},
                new SelectListItem(){Text="Night",Value="N"},
            };

            ViewBag.SewingTeam = new List<SelectListItem>();

            if (setting.SelectedSewingTeam != null)
            {
                List<SelectListItem> SewingTeam = new List<SelectListItem>();
                foreach (var item in setting.SelectedSewingTeam)
                {
                    SewingTeam.Add(new SelectListItem() { Text = item.SewingTeamID, Value = item.SewingTeamID });
                }
                ViewBag.SewingTeam = SewingTeam;
            }

            // 預設值
            //setting.Team = "A";

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
        public ActionResult AQL_AJAX(string AQLPlan, int TotalAvailableQty)
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
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "1" && o.LotSize_Start <= TotalAvailableQty && TotalAvailableQty <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    break;
                case "1.0 Level II":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "2" && o.LotSize_Start <= TotalAvailableQty && TotalAvailableQty <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    break;
                case "1.5 Level I":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == Convert.ToDecimal(1.5) && o.InspectionLevels == "1" && o.LotSize_Start <= TotalAvailableQty && TotalAvailableQty <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    break;
                case "2.5 Level I":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == Convert.ToDecimal(2.5) && o.InspectionLevels == "1" && o.LotSize_Start <= TotalAvailableQty && TotalAvailableQty <= o.LotSize_End).FirstOrDefault();
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
            FinalInspectionService fsevice = new FinalInspectionService();
            FinalInspectionSettingService sevice = new FinalInspectionSettingService();

            string finalInspectionID = string.Empty;

            ViewBag.InspectionStageList = new List<SelectListItem>()
                {
                    new SelectListItem(){Text="Inline",Value="Inline"},
                    new SelectListItem(){Text="Stagger",Value="Stagger"},
                    new SelectListItem(){Text="Final",Value="Final"},
                    new SelectListItem(){Text="3rd Party",Value="3rd Party"},
                };

            ViewBag.Shift = new List<SelectListItem>()
            {
                new SelectListItem(){Text="Day",Value="D"},
                new SelectListItem(){Text="Night",Value="N"},
            };

            ViewBag.SewingTeam = new List<SelectListItem>();

            if (setting.SelectedSewingTeam != null)
            {
                List<SelectListItem> SewingTeam = new List<SelectListItem>();
                foreach (var item in setting.SelectedSewingTeam)
                {
                    SewingTeam.Add(new SelectListItem() { Text = item.SewingTeamID, Value = item.SewingTeamID });
                }
                ViewBag.SewingTeam = SewingTeam;
            }


            ViewBag.AQLPlanList = new List<SelectListItem>()
                {
                    new SelectListItem(){Text="",Value=""},
                    new SelectListItem(){Text="1.0 Level I",Value="1.0 Level I"},
                    new SelectListItem(){Text="1.0 Level II",Value="1.0 Level II"},
                    new SelectListItem(){Text="1.5 Level I",Value="1.5 Level I"},
                    new SelectListItem(){Text="2.5 Level I",Value="2.5 Level I"},
                    new SelectListItem(){Text="100% Inspection",Value="100% Inspection"},
                };


            if (!setting.AcceptQty.HasValue)
            {
                setting.AcceptQty = 0;
            }

            if (setting.SelectedPO != null && setting.SelectedPO.Where(o => string.IsNullOrEmpty(o.Seq)).Any())
            {
                setting.ErrorMessage = $@"
msg.WithError(""Shipmode Seq cant't be empty."");
";
                return View("Setting", setting);
            }

            // Setting存檔，並取得 finalInspectionID
            BaseResult result = sevice.UpdateFinalInspection(setting, this.UserID, this.FactoryID, this.MDivisionID, out finalInspectionID);

            // 錯誤回到Setting頁
            if (!result)
            {
                setting.ErrorMessage = result.ErrorMessage;
                FinalInspectionSettingService finalInspectionSettingService = new FinalInspectionSettingService();
                setting = finalInspectionSettingService.GetSettingForInspection(setting.SelectedPO[0].CustPONO, setting.SelectedPO.Select(o => o.OrderID).ToList(), this.FactoryID);

                if (setting.SelectedSewingTeam != null)
                {
                    List<SelectListItem> SewingTeam = new List<SelectListItem>();
                    foreach (var item in setting.SelectedSewingTeam)
                    {
                        SewingTeam.Add(new SelectListItem() { Text = item.SewingTeamID, Value = item.SewingTeamID });
                    }
                    ViewBag.SewingTeam = SewingTeam;
                }
                // 預設值
                setting.Team = "A";

                return View("Setting", setting);
            }

            fsevice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
            {
                ID = finalInspectionID,
                InspectionStep = "Insp-General"
            }, "Setting", this.UserID);

            return RedirectToAction("General", new { FinalInspectionID = finalInspectionID });
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
            FinalInspectionService sevice = new FinalInspectionService();
            string FinalInspectionID = model.ID;

            if (goPage == "Back")
            {
                model.InspectionStep = "Setting";
                sevice.UpdateFinalInspectionByStep(model, "Insp-General", this.UserID);
                return RedirectToAction("Setting", new { finalInspectionID = FinalInspectionID });
            }
            else if (goPage == "Next")
            {
                model.InspectionStep = "Insp-CheckList";
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
            FinalInspectionService sevice = new FinalInspectionService();
            if (goPage == "Back")
            {
                finalinspection.InspectionStep = "Insp-General";
                sevice.UpdateFinalInspectionByStep(finalinspection, "Insp-CheckList", this.UserID);

                return RedirectToAction("General", new { FinalInspectionID = finalinspection.ID });
            }
            else if (goPage == "Next")
            {
                finalinspection.InspectionStep = "Insp-Measurement";
                sevice.UpdateFinalInspectionByStep(finalinspection, "Insp-CheckList", this.UserID);

                return RedirectToAction("Measurement", new { FinalInspectionID = finalinspection.ID });
            }
            return View();
        }
        #endregion

        #region Measurement頁面
        public ActionResult Measurement(string FinalInspectionID)
        {
            FinalInspectionMeasurementService service = new FinalInspectionMeasurementService();
            Measurement model = service.GetMeasurementForInspection(FinalInspectionID, this.UserID);

            List<string> listSize = model.ListSize.Select(O => O.SizeCode).Distinct().ToList();

            List<SelectListItem> sizeList = new SetListItem().ItemListBinding(listSize);
            ViewBag.ListSize = sizeList;

            TempData["AllSize"] = model.ListSize;

            return View(model);
        }

        /// <summary>
        /// 右上角Save按鈕
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult MeasurementSingleSave(Measurement model)
        {
            if (model.ListMeasurementItem != null && !model.ListMeasurementItem.Where(o => o.ResultSizeSpec != null && o.ResultSizeSpec != "").Any())
            {
                return Json(true);
            }

            FinalInspectionMeasurementService service = new FinalInspectionMeasurementService();
            BaseResult result = service.UpdateMeasurement(model, this.UserID);

            if (result)
            {
                FinalInspectionService fservice = new FinalInspectionService();

                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = model.FinalInspectionID,
                    InspectionStep = "Insp-AddDefect"
                }, "Insp-Measurement", this.UserID);
            }

            return Json(result);
        }

        [HttpPost]
        public ActionResult GetNewSizeByArticle(string Article)
        {
            List<ArticleSize> ListSize = (List<ArticleSize>)TempData["AllSize"];

            var listSize = ListSize.Where(o => o.Article == Article).Select(O => O.SizeCode).Distinct().ToList();
            List<SelectListItem> result = new SetListItem().ItemListBinding(listSize);

            // 保存原資料
            TempData["AllSize"] = ListSize;

            return Json(result);
        }

        [HttpPost]
        public ActionResult OpenView(string FinalInspectionID)
        {
            FinalInspectionMeasurementService service = new FinalInspectionMeasurementService();
            List<MeasurementViewItem> result = service.GetMeasurementViewItem(FinalInspectionID);

            //string someJson = "[  {    \"Code\": \"S021\",    \"Description\": \"A Chest width ( meas. 2cm below armhole )\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"103\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S03\",    \"Description\": \"B WAIST WIDTH\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"99\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S04\",    \"Description\": \"B1 WAIST MEAS. POINT FROM CHEST MEAS. POINT\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"19\",    \"2021/08/20 16:25:38\": \"1\",    \"diff1\": \"-18\"  },  {    \"Code\": \"S05\",    \"Description\": \"D HEM OPENING  (MEAS. STRAIGHT)\",    \"Tol(+)\": \"2\",    \"Tol(-)\": \"1\",    \"50_aa\": \"102\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S06\",    \"Description\": \"E Front zip length (zipend 0,5cm before collar edge) (tolerance +/- 1%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"26\",    \"2021/08/20 16:25:38\": \"2\",    \"diff1\": \"-24\"  },  {    \"Code\": \"S07\",    \"Description\": \"F FRONT TO BACK\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"3\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S08\",    \"Description\": \"H 1/2 ZIP - Shoulder length\",    \"Tol(+)\": \"0.8\",    \"Tol(-)\": \"0.4\",    \"50_aa\": \"14.6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S09\",    \"Description\": \"I 2 PIECE - PRESHAPE  sleeve - Sleeve length\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"67\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S10\",    \"Description\": \"J Sleeve width ( meas. 2cm below armhole )\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"39.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S11\",    \"Description\": \"K LONG + PRESHAPE SLEEVE -Ellbow width (meas. 32,0cm above sleeve opening)\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"28.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S14\",    \"Description\": \"M LONG + PRESHAPE sleeve - Sleeve opening\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"21\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S15\",    \"Description\": \"O1 1/2 ZIP - Front neck drop (HPS to c.f. neck seam)\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0\",    \"50_aa\": \"10.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S16\",    \"Description\": \"O2 1/2 ZIP - Back neck drop (HPS to c.b. neck seam) pattern meas.\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0\",    \"50_aa\": \"2\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S17\",    \"Description\": \"N 1/2 ZIP - Back neck width (HPS to HPS)\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"14.4\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S18\",    \"Description\": \"Q1 COLLAR LENGTH OUTER EDGE\",    \"Tol(+)\": \"1\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"41.6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S19\",    \"Description\": \"Q2 Collar height (center front) (tolerance +/- 10%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"6\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S20\",    \"Description\": \"Q3 COLLAR HEIGHT (CENTER BACK) (TOLERANCE +/- 10%)\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"4.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S21\",    \"Description\": \"S MINIMUM NECK OPENING STRETCHED\",    \"Tol(+)\": \"0\",    \"Tol(-)\": \"0\",    \"50_aa\": \"63\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S22\",    \"Description\": \"Y1 Logo meas.: top edge of logo meas. to HPS\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"17.5\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S23\",    \"Description\": \"Y2 LOGO MEAS.: EDGE OF LOGO TO CENTER FRONT\",    \"Tol(+)\": \"0.5\",    \"Tol(-)\": \"0.5\",    \"50_aa\": \"7\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S24\",    \"Description\": \"Z 1/2 ZIP - Back length\",    \"Tol(+)\": \"1.5\",    \"Tol(-)\": \"1\",    \"50_aa\": \"74\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S25\",    \"Description\": \"DEC LABEL\",    \"Tol(+)\": \"\",    \"Tol(-)\": \"\",    \"50_aa\": \"50\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  },  {    \"Code\": \"S27\",    \"Description\": \"POLYBAG SIZE (WXL)\",    \"Tol(+)\": \"\",    \"Tol(-)\": \"\",    \"50_aa\": \"30X40\",    \"2021/08/20 16:25:38\": null,    \"diff1\": null  }]";

            return Json(result);
        }

        [HttpPost]
        public ActionResult Measurement(Measurement model, string goPage)
        {
            // model的ListMeasurementItem一定會是空的

            FinalInspectionService fservice = new FinalInspectionService();

            if (goPage == "Back")
            {
                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = model.FinalInspectionID,
                    InspectionStep = "Insp-CheckList"
                }, "Insp-Measurement", this.UserID);
                return RedirectToAction("CheckList", new { FinalInspectionID = model.FinalInspectionID });
            }
            else if (goPage == "Next")
            {

                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = model.FinalInspectionID,
                    InspectionStep = "Insp-AddDefect"
                }, "Insp-Measurement", this.UserID);

                return RedirectToAction("AddDefect", new { FinalInspectionID = model.FinalInspectionID });
            }

            return View();
        }

        #endregion

        #region AddDefect頁面
        //暫存圖片用
        public static List<FinalInspectionDefectItem> TmpFinalInspectionDefectItem_List;

        public ActionResult AddDefect(string FinalInspectionID)
        {
            TmpFinalInspectionDefectItem_List = new List<FinalInspectionDefectItem>();
            FinalInspectionSettingService SettingService = new FinalInspectionSettingService();
            FinalInspectionAddDefectService Addsevice = new FinalInspectionAddDefectService();

            Setting model = SettingService.GetSettingForInspection(FinalInspectionID);

            DatabaseObject.ViewModel.FinalInspection.AddDefect addDefct = new DatabaseObject.ViewModel.FinalInspection.AddDefect();
            addDefct = Addsevice.GetDefectForInspection(FinalInspectionID);

            addDefct.FinalInspectionID = FinalInspectionID;

            addDefct.SampleSize = model.SampleSize;

            return View(addDefct);
        }

        public ActionResult DefectPicture(string DetailUkey, string RowIndex)
        {
            DetailUkey = DetailUkey == null || DetailUkey == string.Empty ? "0" : DetailUkey;
            long FinalInspection_DetailUkey = Convert.ToInt64(DetailUkey);
            long FinalInspection_RowIndex = Convert.ToInt64(RowIndex);
            FinalInspectionAddDefectService Addsevice = new FinalInspectionAddDefectService();

            //取得該Detail在DB現有的圖片
            List<byte[]> model = Addsevice.GetDefectImage(FinalInspection_DetailUkey);

            // 把畫面上User拍的照片加進去，一起顯示
            if (TmpFinalInspectionDefectItem_List != null)
            {
                // DB有的就用Ukey
                if (FinalInspection_DetailUkey != 0)
                {
                    foreach (var item in TmpFinalInspectionDefectItem_List.Where(o => o.Ukey == FinalInspection_DetailUkey))
                    {
                        model.Add(item.TempImage);
                    }
                }
                else
                {
                    // DB沒有的就用RowIndex
                    foreach (var item in TmpFinalInspectionDefectItem_List.Where(o => o.RowIndex == FinalInspection_RowIndex))
                    {
                        model.Add(item.TempImage);
                    }
                }
            }

            return View("FinalInspection_Picture", model);
        }

        public ActionResult AddDefectTakePicture(string DetailUkey, string RowIndex)
        {
            ViewData["DetailUkey"] = DetailUkey;
            ViewData["RowIndex"] = RowIndex;
            return View("TakePicture");
        }

        [HttpPost]
        public ActionResult AddDefectPicturesTempSave(FinalInspectionDefectItem data)
        {
            TmpFinalInspectionDefectItem_List.Add(data);
            return Json(true);
        }

        [HttpPost]
        public ActionResult AddDefect(AddDefect addDefct, string goPage)
        {

            FinalInspectionService fservice = new FinalInspectionService();
            AddDefect latestModel = new AddDefect();
            UpdateModel(latestModel);

            addDefct.FinalInspectionID = latestModel.FinalInspectionID;
            addDefct.RejectQty = latestModel.RejectQty;
            addDefct.SampleSize = latestModel.SampleSize;

            if (goPage == "Back")
            {
                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = addDefct.FinalInspectionID,
                    InspectionStep = "Insp-Measurement"
                }, "Insp-AddDefect", this.UserID);

                return RedirectToAction("Measurement", new { FinalInspectionID = addDefct.FinalInspectionID });
            }
            else if (goPage == "Next")
            {
                addDefct.RejectQty = addDefct.RejectQty.HasValue ? addDefct.RejectQty : 0;
                // 本次新增的圖片全面加入
                foreach (var item in addDefct.ListFinalInspectionDefectItem)
                {
                    long FinalInspection_DetailUkey = item.Ukey;
                    long RowIndex = item.RowIndex;

                    if (item.Ukey > 0)
                    {
                        var sameUkeyImg = TmpFinalInspectionDefectItem_List.Where(o => o.Ukey == FinalInspection_DetailUkey);
                        foreach (var data in sameUkeyImg)
                        {
                            item.ListFinalInspectionDefectImage.Add(data.TempImage);
                        }
                    }
                    else
                    {
                        var sameUkeyImg = TmpFinalInspectionDefectItem_List.Where(o => o.RowIndex == RowIndex);
                        foreach (var data in sameUkeyImg)
                        {
                            item.ListFinalInspectionDefectImage.Add(data.TempImage);
                        }
                    }
                }

                FinalInspectionAddDefectService Addsevice = new FinalInspectionAddDefectService();
                Addsevice.UpdateFinalInspectionDetail(addDefct, this.UserID);

                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = addDefct.FinalInspectionID,
                    InspectionStep = "Insp-BA"
                }, "Insp-AddDefect", this.UserID);

                return RedirectToAction("BeautifulProductAudit", new { FinalInspectionID = addDefct.FinalInspectionID });
            }
            return View(addDefct);

        }

        #endregion

        #region Beautiful Product Audit頁面
        public static List<BACriteriaItem> TmpBACriteriaItem_List;

        public ActionResult BeautifulProductAudit(string FinalInspectionID)
        {
            TmpBACriteriaItem_List = new List<BACriteriaItem>();
            FinalInspectionBeautifulProductAuditService Service = new FinalInspectionBeautifulProductAuditService();

            BeautifulProductAudit model = new BeautifulProductAudit() { ListBACriteria = new List<BACriteriaItem>() };
            model = Service.GetBeautifulProductAuditForInspection(FinalInspectionID);

            return View(model);
        }

        public ActionResult BATakePicture(string DetailUkey, string RowIndex)
        {
            ViewData["DetailUkey"] = DetailUkey;
            ViewData["RowIndex"] = RowIndex;
            return View("TakePicture");
        }

        public ActionResult BAPicture(string DetailUkey, string RowIndex)
        {
            DetailUkey = DetailUkey == null || DetailUkey == string.Empty ? "0" : DetailUkey;
            long FinalInspection_DetailUkey = Convert.ToInt64(DetailUkey);
            long FinalInspection_RowIndex = Convert.ToInt64(RowIndex);

            FinalInspectionBeautifulProductAuditService Service = new FinalInspectionBeautifulProductAuditService();


            //取得該Detail在DB現有的圖片
            List<byte[]> model = Service.GetBACriteriaImage(FinalInspection_DetailUkey);

            // 把畫面上User拍的照片加進去，一起顯示
            if (TmpBACriteriaItem_List != null)
            {
                // DB有的就用Ukey
                if (FinalInspection_DetailUkey != 0)
                {
                    foreach (var item in TmpBACriteriaItem_List.Where(o => o.Ukey == FinalInspection_DetailUkey))
                    {
                        model.Add(item.TempImage);
                    }
                }
                else
                {
                    // DB沒有的就用RowIndex
                    foreach (var item in TmpBACriteriaItem_List.Where(o => o.RowIndex == FinalInspection_RowIndex))
                    {
                        model.Add(item.TempImage);
                    }
                }
            }

            return View("FinalInspection_Picture", model);
        }

        [HttpPost]
        public ActionResult AddBaPicturesTempSave(BACriteriaItem data)
        {
            TmpBACriteriaItem_List.Add(data);
            return Json(true);
        }

        [HttpPost]
        public ActionResult BeautifulProductAudit(BeautifulProductAudit Req, string goPage)
        {
            FinalInspectionService fservice = new FinalInspectionService();
            BeautifulProductAudit latestModel = new BeautifulProductAudit();
            UpdateModel(latestModel);

            Req.BAQty = latestModel.BAQty;
            Req.FinalInspectionID = latestModel.FinalInspectionID;

            if (goPage == "Back")
            {
                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = latestModel.FinalInspectionID,
                    InspectionStep = "Insp-AddDefect"
                }, "Insp-BA", this.UserID);

                return RedirectToAction("AddDefect", new { FinalInspectionID = Req.FinalInspectionID });
            }
            else if (goPage == "Next")
            {
                Req.BAQty = Req.BAQty.HasValue ? Req.BAQty.Value : 0;
                // 本次新增的圖片全面加入
                foreach (var item in Req.ListBACriteria)
                {
                    long FinalInspection_DetailUkey = item.Ukey;
                    long RowIndex = item.RowIndex;

                    if (item.Ukey > 0)
                    {
                        var sameUkeyImg = TmpBACriteriaItem_List.Where(o => o.Ukey == FinalInspection_DetailUkey);
                        foreach (var data in sameUkeyImg)
                        {
                            item.ListBACriteriaImage.Add(data.TempImage);
                        }
                    }
                    else
                    {
                        var sameUkeyImg = TmpBACriteriaItem_List.Where(o => o.RowIndex == RowIndex);
                        foreach (var data in sameUkeyImg)
                        {
                            item.ListBACriteriaImage.Add(data.TempImage);
                        }
                    }
                }

                FinalInspectionBeautifulProductAuditService Service = new FinalInspectionBeautifulProductAuditService();
                Service.UpdateBeautifulProductAudit(Req, this.UserID);

                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = latestModel.FinalInspectionID,
                    InspectionStep = "Insp-Moisture"
                }, "Insp-BA", this.UserID);

                return RedirectToAction("Moisture", new { FinalInspectionID = Req.FinalInspectionID });
            }
            return View(Req);

        }

        #endregion

        #region Moisture頁面
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

        public ActionResult Moisture(string FinalInspectionID)
        {
            DatabaseObject.ViewModel.FinalInspection.Moisture model = new DatabaseObject.ViewModel.FinalInspection.Moisture();
            FinalInspectionMoistureService service = new FinalInspectionMoistureService();

            model = service.GetMoistureForInspection(FinalInspectionID);

            ViewBag.ListArticle = new SetListItem().ItemListBinding(model.ListArticle);
            ViewBag.ListCartonItem = model.ListCartonItem;
            ViewBag.ListEndlineMoisture = model.ListEndlineMoisture;
            ViewBag.ActionSelectListItem = model.ActionSelectListItem;

            ViewBag.FinalInspectionID = model.FinalInspectionID;
            ViewBag.FinalInspection_CTNMoistureStandard = model.FinalInspection_CTNMoistureStandard;
            ViewBag.FinalInspection_CTNMoistureStandardBM = model.FinalInspection_CTNMoistureStandardBM;

            DatabaseObject.ViewModel.FinalInspection.MoistureResult moistureResult = new DatabaseObject.ViewModel.FinalInspection.MoistureResult();

            return View(moistureResult);
        }

        [HttpPost]
        public ActionResult GetViewMoistureResult(string finalInspectionID)
        {
            IFinalInspectionMoistureService finalInspectionMoistureService = new FinalInspectionMoistureService();
            List<ViewMoistureResult> viewMoistureResultsList = finalInspectionMoistureService.GetViewMoistureResult(finalInspectionID);

            return Json(viewMoistureResultsList);
        }

        [HttpPost]
        public ActionResult MoistureSingleSave(MoistureResult Req)
        {
            FinalInspectionMoistureService service = new FinalInspectionMoistureService();

            if (Req.Result != null && Req.Result.ToUpper() == "PASS")
            {
                Req.Result = "P";
            }
            else
            {
                Req.Result = "F";
            }

            BaseResult result = service.UpdateMoistureBySave(Req);

            if (result)
            {
                FinalInspectionService fservice = new FinalInspectionService();

                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = Req.FinalInspectionID,
                    InspectionStep = "Insp-Moisture"
                }, "Insp-Moisture", this.UserID);
            }

            return Json(result);
        }

        [HttpPost]
        public ActionResult MoistureDelete(long UKey)
        {
            FinalInspectionMoistureService service = new FinalInspectionMoistureService();
            BaseResult result = service.DeleteMoisture(UKey);

            return Json(result);
        }

        [HttpPost]
        public ActionResult Moisture(DatabaseObject.ViewModel.FinalInspection.MoistureResult moistureResult, string goPage)
        {
            FinalInspectionService fservice = new FinalInspectionService();
            if (goPage == "Back")
            {
                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = moistureResult.FinalInspectionID,
                    InspectionStep = "Insp-BA"
                }, "Insp-Moisture", this.UserID);

                return RedirectToAction("BeautifulProductAudit", new { FinalInspectionID = moistureResult.FinalInspectionID });
            }
            else if (goPage == "Next")
            {
                FinalInspectionMoistureService service = new FinalInspectionMoistureService();
                var MoistureResult = service.UpdateMoistureByNext(moistureResult);

                // 完成資料才可以進入下一步
                if (MoistureResult.Result)
                {
                    fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                    {
                        ID = moistureResult.FinalInspectionID,
                        InspectionStep = "Insp-Others"
                    }, "Insp-Moisture", this.UserID);

                    return RedirectToAction("Others", new { FinalInspectionID = moistureResult.FinalInspectionID });
                }
                else
                {
                    ViewData["ErrorMessage"] = $@"
msg.WithInfo('{MoistureResult.ErrorMessage}');
";

                    Moisture model = service.GetMoistureForInspection(moistureResult.FinalInspectionID);

                    ViewBag.ListArticle = new SetListItem().ItemListBinding(model.ListArticle);
                    ViewBag.ListCartonItem = model.ListCartonItem;
                    ViewBag.ListEndlineMoisture = model.ListEndlineMoisture;
                    ViewBag.ActionSelectListItem = model.ActionSelectListItem;

                    ViewBag.FinalInspectionID = model.FinalInspectionID;
                    ViewBag.FinalInspection_CTNMoistureStandard = model.FinalInspection_CTNMoistureStandard;
                }
            }
            return View(moistureResult);
        }

        #endregion

        #region Others頁面

        public static List<OtherImage> TmpListOthersImageItem_List;
        public ActionResult Others(string FinalInspectionID)
        {
            TmpListOthersImageItem_List = new List<OtherImage>();
            FinalInspectionOthersService service = new FinalInspectionOthersService();
            Others model = new Others();

            model = service.GetOthersForInspection(FinalInspectionID);
            model.CFA = this.UserID;
            return View(model);
        }

        public ActionResult OthersPicture(string FinalInspectionID)
        {
            FinalInspectionOthersService Service = new FinalInspectionOthersService();

            // 取得該DB現有的圖片
            List<OtherImage> list = Service.GetOthersImage(FinalInspectionID);
            List<byte[]> model = list.Select(o => o.Image).ToList();

            // 把畫面上User拍的照片加進去，一起顯示
            if (TmpListOthersImageItem_List != null)
            {
                foreach (var item in TmpListOthersImageItem_List)
                {
                    model.Add(item.TempImage);
                }
            }

            return View("FinalInspection_Picture", model);
        }

        public ActionResult OthersTakePicture(string FinalInspectionID)
        {
            ViewData["DetailUkey"] = FinalInspectionID;

            return View("Others_TakePicture");
        }

        [HttpPost]
        public ActionResult AddthersTPicturesTempSave(OtherImage data)
        {
            TmpListOthersImageItem_List.Add(data);
            return Json(true);
        }

        [HttpPost]
        public ActionResult Others(Others model, string goPage)
        {
            FinalInspectionService fservice = new FinalInspectionService();
            FinalInspectionOthersService oService = new FinalInspectionOthersService();
            QueryService Qservice = new QueryService();

            if (goPage == "Back")
            {
                oService.UpdateOthersBack(model, this.UserID);

                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = model.FinalInspectionID,
                    InspectionStep = "Insp-Moisture"
                }, "Insp-Others", this.UserID);

                return RedirectToAction("Moisture", new { FinalInspectionID = model.FinalInspectionID });
            }
            else if (goPage == "Submit")
            {

                // 本次新增的圖片全面加入
                model.ListOthersImageItem = new List<OtherImage>();
                foreach (var item in TmpListOthersImageItem_List)
                {
                    OtherImage o = new OtherImage();
                    o.ID = model.FinalInspectionID;
                    o.Image = item.TempImage;
                    model.ListOthersImageItem.Add(o);
                }

                fservice.UpdateFinalInspectionByStep(new DatabaseObject.ManufacturingExecutionDB.FinalInspection()
                {
                    ID = model.FinalInspectionID,
                    InspectionStep = "Submit"
                }, "Insp-Others", this.UserID);

                // Submit 紀錄
                BaseResult r = oService.UpdateOthersSubmit(model, this.UserID);

                // 取出剛剛的紀錄
                FinalInspectionService sevice = new FinalInspectionService();
                DatabaseObject.ManufacturingExecutionDB.FinalInspection current = sevice.GetFinalInspection(model.FinalInspectionID);

                // 若是Submit成功，且Fail則寄信
                if (r.Result && current.InspectionResult == "Fail")
                {
                    bool test = IsTest.ToLower() == "true";

                    string WebHost = Request.Url.Scheme + @"://" + Request.Url.Authority + "/";
                    Qservice.SendMail(model.FinalInspectionID, WebHost, test);
                }

                if (r.Result)
                {
                    model.ErrorMessage = @"msg.WithSuccesCheck(""Submit Success, redirect to top page."",function() { window.location.href = ""/FinalInspection/Inspection""; });";
                }
                else
                {
                    model.ErrorMessage = $@"msg.WithError(""Submit Fail, {r.ErrorMessage}"");
";
                }
            }

            return View(model);
        }
        #endregion

    }
}