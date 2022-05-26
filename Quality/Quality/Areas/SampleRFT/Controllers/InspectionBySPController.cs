using BusinessLogicLayer.Service.SampleRFT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.SampleRFT;
using FactoryDashBoardWeb.Helper;
using Newtonsoft.Json;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.SampleRFT.Controllers
{
    public class InspectionBySPController : BaseController
    {
        private InspectionBySPService _Service;
        public InspectionBySPController()
        {
            this.SelectedMenu = "Sample RFT";
            ViewBag.OnlineHelp = this.OnlineHelp + "SampleRFT.InspectionBySP,,";
            _Service = new InspectionBySPService();
        }

        #region 查詢SP#
        public ActionResult Index()
        {
            InspectionBySP_ViewModel model = new InspectionBySP_ViewModel() { DataList = new List<InspectionBySP_SearchResult>() };
            return View(model);
        }

        [HttpPost]
        public ActionResult Index(InspectionBySP_ViewModel Req)
        {
            this.CheckSession();

            InspectionBySP_ViewModel model = new InspectionBySP_ViewModel() { DataList = new List<InspectionBySP_SearchResult>() };
            if (string.IsNullOrEmpty(Req.OrderID) && string.IsNullOrEmpty(Req.CustPONo) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.SeasonID))
            {
                model.ErrorMessage = $@"msg.WithError('Search condition cannot be empty.');";
            }
            else
            {
                model = _Service.Get_SearchResults(Req);

                if (!model.Result)
                {
                    model.ErrorMessage = $@"msg.WithError(""{model.ErrorMessage}"");";
                }
            }

            return View(model);
        }

        #endregion

        #region Setting頁

        public ActionResult Setting(string OrderID)
        {
            this.CheckSession();

            InspectionBySP_Setting model = new InspectionBySP_Setting();
            SampleRFTInspection InProcess = new SampleRFTInspection();

            try
            {
                string InspectionProgress = _Service.CheckInspection(new InspectionBySP_ViewModel() { OrderID = OrderID });

                switch (InspectionProgress)
                {
                    case "NonInspect":
                        model = _Service.GetNewSetting(OrderID, this.FactoryID);
                        break;
                    case "InProcess":
                        // 取得進行到一半的檢驗單後再導過去
                        InProcess = _Service.GetSampleRFTInspections(new InspectionBySP_ViewModel() { OrderID = OrderID }).FirstOrDefault();
                        break;
                    case "Failure":
                        model = _Service.GetNewSetting(OrderID, this.FactoryID);
                        break;

                    default:
                        break;
                }

                if (InspectionProgress == "InProcess" && InProcess.InspectionStep == "Insp-Setting")
                {
                    // 如果是停在Setting，則抓出已存在的資料
                    model = _Service.GetExistedSetting(new InspectionBySP_ViewModel() { ID = InProcess.ID, OrderID = InProcess.OrderID });
                }
                else if (InspectionProgress == "InProcess")
                {
                    // 導向各功能的頁面
                    switch (InProcess.InspectionStep)
                    {
                        case "Insp-CheckList":
                            return RedirectToAction("CheckList", new { ID = InProcess.ID });
                        case "Insp-Measurement":
                            return RedirectToAction("Measurement", new { ID = InProcess.ID });
                        case "Insp-AddDefect":
                            return RedirectToAction("AddDefect", new { ID = InProcess.ID });
                        case "Insp-BA":
                            return RedirectToAction("BeautifulProductAudit", new { ID = InProcess.ID });
                        case "Insp-DummyFit":
                            return RedirectToAction("DummyFitting", new { ID = InProcess.ID });
                        case "Insp-Others":
                            return RedirectToAction("Others", new { ID = InProcess.ID });

                        default:
                            break;
                    }
                }

                if (!model.ExecuteResult)
                {
                    throw new Exception(model.ErrorMessage);
                }

                TempData["Setting"] = model;
                ViewData["RejectQty"] = model.AcceptQty + 1;
            }
            catch (Exception ex)
            {
                model.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }

            return View(model);
        }

        [HttpPost]
        public ActionResult AQL_AJAX(string AQLPlan, int OrderQty)
        {
            this.CheckSession();

            InspectionBySP_Setting setting = new InspectionBySP_Setting();
            if (TempData["Setting"] != null)
            {
                setting = (InspectionBySP_Setting)TempData["Setting"];
            }
            TempData["Setting"] = setting;
            var SamplePlanQty = 0;
            var AcceptedQty = 0;
            var RejectQty = 0;
            long AcceptableQualityLevelsUkey = 0;
            AcceptableQualityLevels tmp = new AcceptableQualityLevels();
            switch (AQLPlan)
            {
                case "1.0 Level":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == 1 && o.InspectionLevels == "1" && o.LotSize_Start <= OrderQty && OrderQty <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    AcceptableQualityLevelsUkey = tmp.Ukey;
                    break;
                case "1.5 Level":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == Convert.ToDecimal(1.5) && o.InspectionLevels == "1" && o.LotSize_Start <= OrderQty && OrderQty <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    AcceptableQualityLevelsUkey = tmp.Ukey;
                    break;
                case "2.5 Level":
                    tmp = setting.AcceptableQualityLevels.Where(o => o.AQLType == Convert.ToDecimal(2.5) && o.InspectionLevels == "1" && o.LotSize_Start <= OrderQty && OrderQty <= o.LotSize_End).FirstOrDefault();
                    SamplePlanQty = tmp.SampleSize.Value;
                    AcceptedQty = tmp.AcceptedQty.Value;
                    AcceptableQualityLevelsUkey = tmp.Ukey;
                    break;
                default:
                    break;
            }
            RejectQty = AcceptedQty + 1;

            var jsonObject = new List<object>();
            jsonObject.Add(JsonConvert.SerializeObject(SamplePlanQty));
            jsonObject.Add(JsonConvert.SerializeObject(AcceptedQty));
            jsonObject.Add(JsonConvert.SerializeObject(RejectQty));
            jsonObject.Add(JsonConvert.SerializeObject(AcceptableQualityLevelsUkey));

            return Json(jsonObject);
        }

        [HttpPost]
        public ActionResult Setting(InspectionBySP_Setting Req)
        {
            this.CheckSession();

            InspectionBySP_Setting model = new InspectionBySP_Setting();
            SampleRFTInspection InProcess = new SampleRFTInspection();

            Req.AddName = this.UserID;
            // 資料開始異動，並取得該筆ID
            InspectionBySP_Setting result = _Service.SettingProcess(Req);

            if (!result.ExecuteResult)
            {
                // 異動出錯，則重新執行一次Setting 頁面進入的動作
                // 因為有可能是要新建立、或是現在檢驗異動失敗，都要重新判斷一次，帶出相關下拉選單的資料來源
                string InspectionProgress = _Service.CheckInspection(new InspectionBySP_ViewModel() { OrderID = Req.OrderID });

                switch (InspectionProgress)
                {
                    case "NonInspect":
                        model = _Service.GetNewSetting(Req.OrderID, this.FactoryID);
                        break;
                    case "InProcess":
                        // 取得進行到一半的檢驗單後再導過去
                        InProcess = _Service.GetSampleRFTInspections(new InspectionBySP_ViewModel() { OrderID = Req.OrderID }).FirstOrDefault();
                        break;
                    case "Failure":
                        model = _Service.GetNewSetting(Req.OrderID, this.FactoryID);
                        break;

                    default:
                        break;
                }

                if (InspectionProgress == "InProcess" && InProcess.InspectionStep == "Insp-Setting")
                {
                    // 如果是停在Setting，則抓出已存在的資料
                    model = _Service.GetExistedSetting(new InspectionBySP_ViewModel() { ID = InProcess.ID, OrderID = InProcess.OrderID });
                }

                TempData["Setting"] = model;
                ViewData["RejectQty"] = model.AcceptQty + 1;
                model.ErrorMessage = $@"msg.WithError(""{result.ErrorMessage}"")";

                return View(model);
            }
            else
            {
                // 進度更新
                _Service.UpdateSampleRFTInspectionByStep(new SampleRFTInspection()
                {
                    ID = result.ID,
                    InspectionStep = "Insp-CheckList"
                }, "Insp-Setting", this.UserID);
                return RedirectToAction("CheckList", new { ID = result.ID });

            }
        }
        #endregion

        #region CheckList頁
        public ActionResult CheckList(long ID)
        {
            InspectionBySP_CheckList model = new InspectionBySP_CheckList();
            try
            {
                SampleRFTInspection data = _Service.GetSampleRFTInspections(new InspectionBySP_ViewModel() { ID = ID }).FirstOrDefault();

                model.CheckFabricApproval = data.CheckFabricApproval;
                model.CheckMetalDetection = data.CheckMetalDetection;
                model.CheckSealingSampleApproval = data.CheckSealingSampleApproval;
                model.CheckColorShade = data.CheckColorShade;
                model.CheckHandfeel = data.CheckHandfeel;
                model.CheckAppearance = data.CheckAppearance;
                model.CheckPrintEmbDecorations = data.CheckPrintEmbDecorations;
                model.CheckFiberContent = data.CheckFiberContent;
                model.CheckCareInstructions = data.CheckCareInstructions;
                model.CheckDecorativeLabel = data.CheckDecorativeLabel;
                model.CheckCountryofOrigin = data.CheckCountryofOrigin;
                model.CheckSizeKey = data.CheckSizeKey;
                model.CheckAdditionalLabel = data.CheckAdditionalLabel;
                model.CheckPolytagMarketing = data.CheckPolytagMarketing;
                model.CheckCareLabel = data.CheckCareLabel;
                model.CheckSecurityLabel = data.CheckSecurityLabel;
                model.CheckOuterCarton = data.CheckOuterCarton;
                model.CheckPackingMode = data.CheckPackingMode;
                model.CheckHangtag = data.CheckHangtag;
            }
            catch (Exception ex)
            {
                model.ErrorMessage = $@"msg.WithError(""{ex.Message}"")";
            }

            return View(model);
        }

        [HttpPost]
        public ActionResult CheckList(InspectionBySP_CheckList Req, string goPage)
        {
            this.CheckSession();


            if (goPage == "Back")
            {
                Req.InspectionStep = "Insp-Setting";
                _Service.UpdateSampleRFTInspectionByStep(Req, "Insp-CheckList", this.UserID);

                return RedirectToAction("Setting", new { ID = Req.ID });
            }
            else if (goPage == "Next")
            {
                Req.InspectionStep = "Insp-Measurement";
                _Service.UpdateSampleRFTInspectionByStep(Req, "Insp-CheckList", this.UserID);

                return RedirectToAction("Measurement", new { ID = Req.ID });
            }
            return View();
        }
        #endregion

        #region Measurement頁面

        public static List<SelectListItem> TmpMeasurementImgSourceList;
        public static List<RFT_Inspection_Measurement_Image> TmpAdd_MeasurementImg;
        public static List<RFT_Inspection_Measurement_Image> TmpDelete_MeasurementImg;

        public ActionResult Measurement(long ID)
        {

            TmpMeasurementImgSourceList = new List<SelectListItem>();
            TmpAdd_MeasurementImg = new List<RFT_Inspection_Measurement_Image>();
            TmpDelete_MeasurementImg = new List<RFT_Inspection_Measurement_Image>();

            InspectionBySP_Measurement model = new InspectionBySP_Measurement();
            try
            {
                model = _Service.GetMeasurement(ID, this.UserID);


                List<string> listSize = model.ListSize.Select(O => O.SizeCode).Distinct().ToList();

                List<SelectListItem> sizeList = new SetListItem().ItemListBinding(listSize);
                ViewBag.ListSize = sizeList;

                TempData["AllSize"] = model.ListSize;
            }
            catch (Exception ex)
            {
                model.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }

            return View(model);
        }

        /// <summary>
        /// Photo按鈕點選
        /// </summary>
        /// <param name="OrderID"></param>
        /// <returns></returns>
        public ActionResult MeasurementPicture(string OrderID)
        {

            Measurement_ResultModel model = _Service.GetMeasurementImageList(OrderID);
            // 這裡要帶出所有圖片，可以參考Inspection的Measurement Photo怎麼做

            model.OrderID = OrderID;

            // 把畫面上User拍的照片加進去，一起顯示
            if (TmpAdd_MeasurementImg != null)
            {
                int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
                int ctn = 0;
                foreach (var item in TmpAdd_MeasurementImg)
                {
                    ctn += 1;
                    string seq = (maxCtn + ctn).ToString();
                    model.Images_Source.Add(new SelectListItem() { Text = seq, Value = string.Empty });
                    model.Images.Add(new RFT_Inspection_Measurement_Image() { Image = item.TempImage, Seq = Convert.ToInt32(seq) });
                }
            }

            // 下拉選單排除要刪除的
            model.Images_Source = model.Images_Source.Where(o =>
                !TmpDelete_MeasurementImg.Any(x => x.Seq.ToString() == o.Text)
            ).ToList();

            // 圖片排除要刪除的
            model.Images = model.Images.Where(o =>
                !TmpDelete_MeasurementImg.Any(x => x.Seq == o.Seq)
            ).ToList();


            return View(model);
        }


        /// <summary>
        /// WebCam拍攝照片
        /// </summary>
        /// <param name="data"></param>
        /// <param name="OrderID"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult AddMeasurementPicTemp(RFT_Inspection_Measurement_Image data, string OrderID)
        {
            if (data.TempImage != null)
            {

                Measurement_ResultModel model = _Service.GetMeasurementImageList(OrderID);
                // 這裡要帶出所有圖片，可以參考Inspection的Measurement Photo怎麼做

                model.OrderID = OrderID;

                // 把畫面上User拍的照片加進去，一起顯示
                if (TmpAdd_MeasurementImg != null)
                {
                    int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
                    int ctn = 0;
                    foreach (var item in TmpAdd_MeasurementImg)
                    {
                        ctn += 1;
                        string seq = (maxCtn + ctn).ToString();
                        model.Images_Source.Add(new SelectListItem() { Text = seq, Value = string.Empty });
                        model.Images.Add(new RFT_Inspection_Measurement_Image() { Image = item.TempImage, Seq = Convert.ToInt32(seq) });
                    }
                }

                string newSeq = (model.Images_Source.Count + 1).ToString();
                data.Seq = Convert.ToInt32(newSeq);
                TmpAdd_MeasurementImg.Add(data);
            }

            JsonResult json = MeasurementImageRefresh(OrderID);
            return json;
        }

        /// <summary>
        /// 檔案上傳
        /// </summary>
        /// <param name="list"></param>
        /// <param name="OrderID"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult BatchAddMeasurementPicTemp(List<RFT_Inspection_Measurement_Image> list, string OrderID)
        {
            if (list != null && list.Any())
            {
                Measurement_ResultModel model = _Service.GetMeasurementImageList(OrderID);
                // 這裡要帶出所有圖片，可以參考Inspection的Measurement Photo怎麼做

                model.OrderID = OrderID;

                // 把畫面上User拍的照片加進去，一起顯示
                if (TmpAdd_MeasurementImg != null)
                {
                    int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
                    int ctn = 0;
                    foreach (var item in TmpAdd_MeasurementImg)
                    {
                        ctn += 1;
                        string seq = (maxCtn + ctn).ToString();
                        model.Images_Source.Add(new SelectListItem() { Text = seq, Value = string.Empty });
                        model.Images.Add(new RFT_Inspection_Measurement_Image() { Image = item.TempImage, Seq = Convert.ToInt32(seq) });
                    }
                }

                int newMaxCtn = model.Images_Source.Count;
                int newCtn = 0;
                foreach (var data in list.Where(o => o.TempImage != null))
                {
                    newCtn += 1;
                    string seq = (newMaxCtn + newCtn).ToString();
                    data.Seq = Convert.ToInt32(seq);
                    TmpAdd_MeasurementImg.Add(data);
                }
            }

            JsonResult json = MeasurementImageRefresh(OrderID);
            return json;
        }

        [HttpPost]
        public ActionResult DeleteMeasurementPic(RFT_Inspection_Measurement_Image data, string OrderID)
        {
            TmpDelete_MeasurementImg.Add(data);

            JsonResult json = MeasurementImageRefresh(OrderID);
            return json;
        }

        /// <summary>
        /// 照片新增/刪除調整
        /// </summary>
        /// <param name="OrderID"></param>
        /// <returns></returns>
        public JsonResult MeasurementImageRefresh(string OrderID)
        {
            // 取得DB現有的圖片
            Measurement_ResultModel model = _Service.GetMeasurementImageList(OrderID);

            // 取得當前圖片下拉選單，最大Seq
            int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
            int ctn = 0;

            // 將上傳/拍攝的圖片，和DB圖片合併，同時給新的Seq
            foreach (var item in TmpAdd_MeasurementImg)
            {
                ctn += 1;
                string seq = (maxCtn + ctn).ToString();
                model.Images_Source.Add(new SelectListItem() { Text = seq, Value = string.Empty });
                model.Images.Add(new RFT_Inspection_Measurement_Image() { Image = item.TempImage, Seq = Convert.ToInt32(seq) });
            }

            // 下拉選單排除要刪除的
            model.Images_Source = model.Images_Source.Where(o =>
                !TmpDelete_MeasurementImg.Any(x => x.Seq.ToString() == o.Text)
            ).ToList();

            // 圖片排除要刪除的
            model.Images = model.Images.Where(o =>
                !TmpDelete_MeasurementImg.Any(x => x.Seq == o.Seq)
            ).ToList();


            var jsonObject = new List<object>();
            jsonObject.Add(JsonConvert.SerializeObject(model.Images_Source));
            jsonObject.Add(JsonConvert.SerializeObject(model.Images.Select(o => new { o.Seq, o.Image, o.ID })));

            return new JsonResult
            {
                Data = jsonObject,
                MaxJsonLength = int.MaxValue,/*重點在這行*/
                JsonRequestBehavior = JsonRequestBehavior.AllowGet
            };
        }

        [HttpPost]
        public ActionResult GetNewSizeByArticle(string Article)
        {
            TmpMeasurementImgSourceList = new List<SelectListItem>();
            TmpAdd_MeasurementImg = new List<RFT_Inspection_Measurement_Image>();
            TmpDelete_MeasurementImg = new List<RFT_Inspection_Measurement_Image>();

            List<ArticleSize> ListSize = (List<ArticleSize>)TempData["AllSize"];

            var listSize = ListSize.Where(o => o.Article == Article).Select(O => O.SizeCode).Distinct().ToList();
            List<SelectListItem> result = new SetListItem().ItemListBinding(listSize);

            // 保存原資料
            TempData["AllSize"] = ListSize;

            return Json(result);
        }

        /// <summary>
        /// 右上角Save按鈕
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult MeasurementSingleSave(InspectionBySP_Measurement Req)
        {
            this.CheckSession();

            if (Req.ListMeasurementItem != null && !Req.ListMeasurementItem.Where(o => o.ResultSizeSpec != null && o.ResultSizeSpec != "").Any())
            {
                return Json(true);
            }


            // 取得DB現有的圖片
            Measurement_ResultModel model = _Service.GetMeasurementImageList(Req.OrderID);

            foreach (var item in TmpAdd_MeasurementImg)
            {
                model.Images.Add(new RFT_Inspection_Measurement_Image() { Image = item.TempImage, Seq = item.Seq });
            }

            // 圖片排除要刪除的
            model.Images = model.Images.Where(o =>
                !TmpDelete_MeasurementImg.Any(x => x.Seq == o.Seq)
            ).ToList();

            Req.ImageList = model.Images.Select(o => o.Image).ToList();

            if (Req.ListMeasurementItem != null && !Req.ListMeasurementItem.Where(o => o.ResultSizeSpec != null && o.ResultSizeSpec != "").Any() && Req.ImageList.Count == 0)
            {
                return Json(true);
            }


            InspectionBySP_Measurement result = _Service.InsertMeasurement(Req);

            TmpMeasurementImgSourceList = new List<SelectListItem>();
            TmpAdd_MeasurementImg = new List<RFT_Inspection_Measurement_Image>();
            TmpDelete_MeasurementImg = new List<RFT_Inspection_Measurement_Image>();

            return Json(result);
        }

        [HttpPost]
        public ActionResult Measurement(InspectionBySP_Measurement Req, string goPage)
        {
            this.CheckSession();


            if (goPage == "Back")
            {
                Req.InspectionStep = "Insp-CheckList";
                _Service.UpdateSampleRFTInspectionByStep(Req, "Insp-Measurement", this.UserID);

                return RedirectToAction("CheckList", new { ID = Req.ID });
            }
            else if (goPage == "Next")
            {
                Req.InspectionStep = "Insp-AddDefect";
                _Service.UpdateSampleRFTInspectionByStep(Req, "Insp-Measurement", this.UserID);

                return RedirectToAction("AddDefect", new { ID = Req.ID });
            }
            return View();
        }
        #endregion

        #region AddDefect頁面
        public ActionResult AddDefect(long ID)
        {
            this.CheckSession();

            TmpDefectImgSourceList = new List<SelectListItem>();
            TmpAdd_DefectImg = new List<DefectImage>();
            TmpDelete_DefectImg = new List<DefectImage>();

            InspectionBySP_AddDefect model = _Service.GetAddDefectBody(ID);


            return View(model);
        }

        /// <summary>
        /// 圖片的下拉選單
        /// </summary>
        public static List<SelectListItem> TmpDefectImgSourceList;

        /// <summary>
        /// 存放新增的圖片
        /// </summary>
        public static List<DefectImage> TmpAdd_DefectImg;

        /// <summary>
        /// 存放刪除的圖片
        /// </summary>
        public static List<DefectImage> TmpDelete_DefectImg;
        /// <summary>
        /// AddDefect圖片，使用DetailKey + GarmentDefectCodeID作為圖片的索引鍵
        /// </summary>
        /// <param name="ID"></param>
        /// <param name="SampleRFTInspection_DetailUKey"></param>
        /// <param name="GarmentDefectCodeID"></param>
        /// <returns></returns>
        public ActionResult AddDefectPicture(long ID, long SampleRFTInspection_DetailUKey, string GarmentDefectCodeID)
        {
            SampleRFTInspection_Summary model = SampleRFTInspection_DetailUKey > 0 ? _Service.GetDefectImageList(ID, SampleRFTInspection_DetailUKey)
                : new SampleRFTInspection_Summary()
                {
                    Images_Source = new List<SelectListItem>(),
                    Images = new List<DefectImage>(),
                };

            model.ID = ID;
            model.UKey = SampleRFTInspection_DetailUKey;
            model.GarmentDefectCodeID = GarmentDefectCodeID;

            // 把畫面上User拍的照片加進去，一起顯示
            if (TmpAdd_DefectImg != null)
            {
                int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
                int ctn = 0;
                foreach (var item in TmpAdd_DefectImg)
                {
                    if (item.GarmentDefectCodeID != GarmentDefectCodeID)
                    {
                        continue;
                    }
                    ctn += 1;
                    string seq = (maxCtn + ctn).ToString();
                    model.Images_Source.Add(new SelectListItem() { Text = seq, Value = GarmentDefectCodeID });
                    model.Images.Add(new DefectImage() { Image = item.TempImage, GarmentDefectCodeID = GarmentDefectCodeID, Seq = Convert.ToInt32(seq) });
                }
            }

            // 下拉選單排除要刪除的
            model.Images_Source = model.Images_Source.Where(o =>
                !TmpDelete_DefectImg.Any(x => x.Seq.ToString() == o.Text && x.GarmentDefectCodeID == o.Value)
            ).ToList();

            // 圖片排除要刪除的
            model.Images = model.Images.Where(o =>
                !TmpDelete_DefectImg.Any(x => x.Seq == o.Seq && x.GarmentDefectCodeID == o.GarmentDefectCodeID)
            ).ToList();


            return View(model);
        }

        /// <summary>
        /// WebCam拍攝照片
        /// </summary>
        /// <param name="data"></param>
        /// <param name="OrderID"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult AddDefectPicTemp(DefectImage data, long ID, long SampleRFTInspection_DetailUKey, string GarmentDefectCodeID)
        {
            if (data.TempImage != null)
            {
                SampleRFTInspection_Summary model = SampleRFTInspection_DetailUKey > 0 ? _Service.GetDefectImageList(ID, SampleRFTInspection_DetailUKey)
                    : new SampleRFTInspection_Summary()
                    {
                        Images_Source = new List<SelectListItem>(),
                        Images = new List<DefectImage>(),
                    };

                model.ID = ID;
                model.UKey = SampleRFTInspection_DetailUKey;
                model.GarmentDefectCodeID = GarmentDefectCodeID;

                // 把畫面上User拍的照片加進去，一起顯示
                if (TmpAdd_DefectImg != null)
                {
                    int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
                    int ctn = 0;
                    foreach (var item in TmpAdd_DefectImg)
                    {
                        if (item.GarmentDefectCodeID != GarmentDefectCodeID)
                        {
                            continue;
                        }
                        ctn += 1;
                        string seq = (maxCtn + ctn).ToString();
                        model.Images_Source.Add(new SelectListItem() { Text = seq, Value = GarmentDefectCodeID });
                    }
                }

                string newSeq = (model.Images_Source.Count + 1).ToString();
                data.Seq = Convert.ToInt32(newSeq);
                data.SampleRFTInspectionDetailUKey = SampleRFTInspection_DetailUKey;
                data.GarmentDefectCodeID = GarmentDefectCodeID;
                TmpAdd_DefectImg.Add(data);
            }

            JsonResult json = DefectImageRefresh(ID, SampleRFTInspection_DetailUKey, GarmentDefectCodeID);
            return json;
        }

        [HttpPost]
        public ActionResult BatchAddDefectPicTemp(List<DefectImage> list, long ID, long SampleRFTInspection_DetailUKey, string GarmentDefectCodeID)
        {
            if (list != null && list.Any())
            {

                int maxCtn = 0;
                int ctn = 0;

                SampleRFTInspection_Summary model = SampleRFTInspection_DetailUKey > 0 ? _Service.GetDefectImageList(ID, SampleRFTInspection_DetailUKey)
                    : new SampleRFTInspection_Summary()
                    {
                        Images_Source = new List<SelectListItem>(),
                        Images = new List<DefectImage>(),
                    };

                model.ID = ID;
                model.UKey = SampleRFTInspection_DetailUKey;
                model.GarmentDefectCodeID = GarmentDefectCodeID;

                // 把畫面上User拍的照片加進去，一起顯示
                if (TmpAdd_DefectImg != null)
                {
                    maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
                    ctn = 0;
                    foreach (var item in TmpAdd_DefectImg)
                    {
                        if (item.GarmentDefectCodeID != GarmentDefectCodeID)
                        {
                            continue;
                        }
                        ctn += 1;
                        string seq = (maxCtn + ctn).ToString();
                        model.Images_Source.Add(new SelectListItem() { Text = seq, Value = GarmentDefectCodeID });
                    }
                }


                maxCtn = model.Images_Source.Count;
                ctn = 0;
                foreach (var data in list.Where(o => o.TempImage != null))
                {
                    ctn += 1;
                    string seq = (maxCtn + ctn).ToString();
                    data.Seq = Convert.ToInt32(seq);
                    data.SampleRFTInspectionDetailUKey = SampleRFTInspection_DetailUKey;
                    data.GarmentDefectCodeID = GarmentDefectCodeID;
                    TmpAdd_DefectImg.Add(data);
                }
            }

            JsonResult json = DefectImageRefresh(ID, SampleRFTInspection_DetailUKey, GarmentDefectCodeID);
            return json;
        }

        [HttpPost]
        public ActionResult DeleteDefectPic(DefectImage data, long ID, long SampleRFTInspection_DetailUKey, string GarmentDefectCodeID)
        {
            data.SampleRFTInspectionDetailUKey = SampleRFTInspection_DetailUKey;
            data.GarmentDefectCodeID = GarmentDefectCodeID;
            TmpDelete_DefectImg.Add(data);

            JsonResult json = DefectImageRefresh(ID, SampleRFTInspection_DetailUKey, GarmentDefectCodeID);
            return json;
        }

        public JsonResult DefectImageRefresh(long ID, long SampleRFTInspection_Detail, string GarmentDefectCodeID)
        {
            // 取得DB現有的圖片
            SampleRFTInspection_Summary model = SampleRFTInspection_Detail > 0 ? _Service.GetDefectImageList(ID, SampleRFTInspection_Detail)
                : new SampleRFTInspection_Summary()
                {
                    Images_Source = new List<SelectListItem>(),
                    Images = new List<DefectImage>(),
                };

            // 取得當前圖片下拉選單，最大Seq
            int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
            int ctn = 0;

            // 將上傳/拍攝的圖片，和DB圖片合併，同時給新的Seq
            foreach (var item in TmpAdd_DefectImg)
            {
                if (item.GarmentDefectCodeID != GarmentDefectCodeID)
                {
                    continue;
                }
                ctn += 1;
                string seq = (maxCtn + ctn).ToString();
                model.Images_Source.Add(new SelectListItem() { Text = seq, Value = GarmentDefectCodeID });
                model.Images.Add(new DefectImage() { Image = item.TempImage, GarmentDefectCodeID = GarmentDefectCodeID, Seq = Convert.ToInt32(seq) });
            }

            // 下拉選單排除要刪除的
            model.Images_Source = model.Images_Source.Where(o =>
                !TmpDelete_DefectImg.Any(x => x.Seq.ToString() == o.Text && x.GarmentDefectCodeID == o.Value)
            ).ToList();

            // 圖片排除要刪除的
            model.Images = model.Images.Where(o =>
                !TmpDelete_DefectImg.Any(x => x.Seq == o.Seq && x.GarmentDefectCodeID == o.GarmentDefectCodeID)
            ).ToList();


            var jsonObject = new List<object>();
            jsonObject.Add(JsonConvert.SerializeObject(model.Images_Source));
            jsonObject.Add(JsonConvert.SerializeObject(model.Images.Select(o => new { o.Seq, o.Image, o.ImageUKey })));

            return new JsonResult
            {
                Data = jsonObject,
                MaxJsonLength = int.MaxValue,/*重點在這行*/
                JsonRequestBehavior = JsonRequestBehavior.AllowGet
            };
        }

        [HttpPost]
        public ActionResult AddDefect(InspectionBySP_AddDefect addDefct, string goPage)
        {
            this.CheckSession();

            InspectionBySP_AddDefect latestModel = new InspectionBySP_AddDefect();
            UpdateModel(latestModel);


            addDefct.RejectQty = latestModel.RejectQty;
            addDefct.ID = latestModel.ID;


            if (goPage == "Back")
            {
                // 進度更新
                _Service.UpdateSampleRFTInspectionByStep(new SampleRFTInspection()
                {
                    ID = addDefct.ID,
                    InspectionStep = "Insp-Measurement"
                }, "Insp-AddDefect", this.UserID);
                return RedirectToAction("Measurement", new { ID = addDefct.ID });
            }

            var DbImage = TmpAdd_DefectImg.Where(o => o.ImageUKey > 0).ToList();
            var NotDbImage = TmpAdd_DefectImg.Where(o => o.ImageUKey <= 0).ToList();

            // Reject Qty > 0的數量代表有表身資料
            // 開始塞入圖片
            foreach (SampleRFTInspection_Summary item in addDefct.ListDefectItem)
            {
                string GarmentDefectCodeID = item.GarmentDefectCodeID;

                // 相同Detail Ukey 且排除刪除的圖片
                List<DefectImage> tempImgs =  NotDbImage.Where(o => o.GarmentDefectCodeID == GarmentDefectCodeID &&
                                                !TmpDelete_DefectImg.Any(x => x.Seq == o.Seq && x.GarmentDefectCodeID == o.GarmentDefectCodeID)
                                            ).ToList();
                if (tempImgs.Any())
                {
                    foreach (var t in tempImgs)
                    {
                        t.Image = t.TempImage;
                    }

                    item.Images = tempImgs;
                }
            }

            // 可以直接執行刪除的圖片
            var DeleteImg = TmpDelete_DefectImg.Where(o => o.ImageUKey > 0).ToList();

            InspectionBySP_AddDefect result =  _Service.AddDefectProcess(addDefct, DeleteImg);

            if (!result.ExecuteResult)
            {
                TmpDefectImgSourceList = new List<SelectListItem>();
                TmpAdd_DefectImg = new List<DefectImage>();
                TmpDelete_DefectImg = new List<DefectImage>();

                InspectionBySP_AddDefect model = _Service.GetAddDefectBody(addDefct.ID);

                model.ErrorMessage = $@"msg.WithError(""{result.ErrorMessage}"")";
                return View(model);

            }
            else
            {
                // 進度更新
                _Service.UpdateSampleRFTInspectionByStep(new SampleRFTInspection()
                {
                    ID = addDefct.ID,
                    InspectionStep = "Insp-BA"
                }, "Insp-AddDefect", this.UserID);
                return RedirectToAction("BeautifulProductAudit", new { ID = addDefct.ID });
            }

            // 只要Detail UKey底下的圖片有異動，則全部刪除，重新INSERT

        }
        #endregion

        #region BA頁面
        public ActionResult BeautifulProductAudit(long ID)
        {
            this.CheckSession();

            TmpBAImgSourceList = new List<SelectListItem>();
            TmpAdd_BAImg = new List<BAImage>();
            TmpDelete_BAImg = new List<BAImage>();

            InspectionBySP_BA model = _Service.GetBA_Body(ID);


            return View(model);
        }
        /// <summary>
        /// 圖片的下拉選單
        /// </summary>
        public static List<SelectListItem> TmpBAImgSourceList;

        /// <summary>
        /// 存放新增的圖片
        /// </summary>
        public static List<BAImage> TmpAdd_BAImg;

        /// <summary>
        /// 存放刪除的圖片
        /// </summary>
        public static List<BAImage> TmpDelete_BAImg;
        public ActionResult BAPicture(long ID, long BAUKey, string BACriteria)
        {
            this.CheckSession();

            BACriteriaItem model = BAUKey > 0 ? _Service.GetBAImageList(ID, BAUKey)
                : new BACriteriaItem()
                {
                    Images_Source = new List<SelectListItem>(),
                    Images = new List<BAImage>(),
                };
            model.ID = ID;
            model.Ukey = BAUKey;
            model.BACriteria = BACriteria;

            // 把畫面上User拍的照片加進去，一起顯示
            if (TmpAdd_BAImg != null)
            {
                int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
                int ctn = 0;
                foreach (var item in TmpAdd_BAImg)
                {
                    if (item.BACriteria != BACriteria)
                    {
                        continue;
                    }

                    ctn += 1;
                    string seq = (maxCtn + ctn).ToString();
                    model.Images_Source.Add(new SelectListItem() { Text = seq, Value = BACriteria });
                    model.Images.Add(new BAImage() { Image = item.TempImage
                        , SampleRFTInspection_NonBACriteriaUKey = item.SampleRFTInspection_NonBACriteriaUKey
                        , BACriteria = BACriteria
                        , Seq = Convert.ToInt32(seq) });
                }
            }
            // 下拉選單排除要刪除的
            model.Images_Source = model.Images_Source.Where(o =>
                !TmpDelete_BAImg.Any(x => x.Seq.ToString() == o.Text && x.BACriteria ==BACriteria)
            ).ToList();

            // 圖片排除要刪除的
            model.Images = model.Images.Where(o =>
                !TmpDelete_BAImg.Any(x => x.Seq == o.Seq && x.BACriteria == BACriteria)
            ).ToList();

            return View(model);
        }

        [HttpPost]
        public ActionResult AddBAPicTemp(BAImage data, long ID, long BAUKey, string BACriteria)
        {
            if (data.TempImage != null)
            {
                BACriteriaItem model = BAUKey > 0 ? _Service.GetBAImageList(ID, BAUKey)
                    : new BACriteriaItem()
                    {
                        Images_Source = new List<SelectListItem>(),
                        Images = new List<BAImage>(),
                    };
                model.ID = ID;
                model.Ukey = BAUKey;
                model.BACriteria = BACriteria;

                // 把畫面上User拍的照片加進去，一起顯示
                if (TmpAdd_BAImg != null)
                {
                    int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
                    int ctn = 0;
                    foreach (var item in TmpAdd_BAImg)
                    {
                        if (item.BACriteria != BACriteria)
                        {
                            continue;
                        }

                        ctn += 1;
                        string seq = (maxCtn + ctn).ToString();
                        model.Images_Source.Add(new SelectListItem() { Text = seq, Value = BACriteria });
                    }
                }

                string newseq = (model.Images_Source.Count + 1).ToString();
                data.Seq = Convert.ToInt32(newseq);
                data.BACriteria = BACriteria;
                data.SampleRFTInspection_NonBACriteriaUKey = BAUKey;
                TmpAdd_BAImg.Add(data);
            }

            JsonResult json = DefectBARefresh(ID, BAUKey, BACriteria);
            return json;
        }
        [HttpPost]
        public ActionResult BatchAddBAPicTemp(List<BAImage> list, long ID, long BAUKey, string BACriteria)
        {
            if (list != null && list.Any())
            {
                int maxCtn = 0;
                int ctn = 0;

                BACriteriaItem model = BAUKey > 0 ? _Service.GetBAImageList(ID, BAUKey)
                    : new BACriteriaItem()
                    {
                        Images_Source = new List<SelectListItem>(),
                        Images = new List<BAImage>(),
                    };
                model.ID = ID;
                model.Ukey = BAUKey;
                model.BACriteria = BACriteria;

                // 把畫面上User拍的照片加進去，一起顯示
                if (TmpAdd_BAImg != null)
                {
                    maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
                    ctn = 0;
                    foreach (var item in TmpAdd_BAImg)
                    {
                        if (item.BACriteria != BACriteria)
                        {
                            continue;
                        }

                        ctn += 1;
                        string seq = (maxCtn + ctn).ToString();
                        model.Images_Source.Add(new SelectListItem() { Text = seq, Value = BACriteria });
                    }
                }

                maxCtn = model.Images_Source.Count;
                ctn = 0;
                foreach (var data in list.Where(o => o.TempImage != null))
                {
                    ctn += 1;
                    string seq = (maxCtn + ctn).ToString();
                    data.Seq = Convert.ToInt32(seq);
                    data.SampleRFTInspection_NonBACriteriaUKey = BAUKey;
                    data.BACriteria = BACriteria;
                    TmpAdd_BAImg.Add(data);
                }
            }

            JsonResult json = DefectBARefresh(ID, BAUKey, BACriteria);
            return json;
        }
        [HttpPost]
        public ActionResult DeleteBAPic(BAImage data, long ID, long BAUkey, string BACriteria)
        {
            data.SampleRFTInspection_NonBACriteriaUKey = BAUkey;
            data.BACriteria = BACriteria;
            TmpDelete_BAImg.Add(data);

            JsonResult json = DefectBARefresh(ID, BAUkey, BACriteria);
            return json;
        }
        public JsonResult DefectBARefresh(long ID, long BAUKey, string BACriteria)
        {
            // 取得DB現有的圖片
            BACriteriaItem model = BAUKey > 0 ? _Service.GetBAImageList(ID, BAUKey)
                : new BACriteriaItem()
                {
                    Images_Source = new List<SelectListItem>(),
                    Images = new List<BAImage>(),
                };

            // 取得當前圖片下拉選單，最大Seq
            int maxCtn = model.Images_Source.Any() ? model.Images_Source.Max(o => Convert.ToInt32(o.Text)) : 0;
            int ctn = 0;

            // 將上傳/拍攝的圖片，和DB圖片合併，同時給新的Seq
            foreach (var item in TmpAdd_BAImg)
            {
                if (item.BACriteria != BACriteria)
                {
                    continue;
                }

                ctn += 1;
                string seq = (maxCtn + ctn).ToString();
                model.Images_Source.Add(new SelectListItem() { Text = seq, Value = seq });
                model.Images.Add(new BAImage() { Image = item.TempImage
                    , SampleRFTInspection_NonBACriteriaUKey = BAUKey
                    , BACriteria = BACriteria
                    , Seq = Convert.ToInt32(seq) });
            }

            // 下拉選單排除要刪除的
            model.Images_Source = model.Images_Source.Where(o =>
                !TmpDelete_BAImg.Any(x => x.Seq.ToString() == o.Text && x.BACriteria == BACriteria)
            ).ToList();

            // 圖片排除要刪除的
            model.Images = model.Images.Where(o =>
                !TmpDelete_BAImg.Any(x => x.Seq == o.Seq && x.BACriteria == BACriteria)
            ).ToList();


            var jsonObject = new List<object>();
            jsonObject.Add(JsonConvert.SerializeObject(model.Images_Source));
            jsonObject.Add(JsonConvert.SerializeObject(model.Images.Select(o => new { o.Seq, o.Image, o.ImageUKey })));

            return new JsonResult
            {
                Data = jsonObject,
                MaxJsonLength = int.MaxValue,/*重點在這行*/
                JsonRequestBehavior = JsonRequestBehavior.AllowGet
            };
        }

        [HttpPost]
        public ActionResult BeautifulProductAudit(InspectionBySP_BA Req, string goPage)
        {
            this.CheckSession();

            InspectionBySP_AddDefect latestModel = new InspectionBySP_AddDefect();
            UpdateModel(latestModel);


            Req.BAQty = latestModel.BAQty;
            Req.ID = latestModel.ID;

            if (goPage == "Back")
            {
                // 進度更新
                _Service.UpdateSampleRFTInspectionByStep(new SampleRFTInspection()
                {
                    ID = Req.ID,
                    InspectionStep = "Insp-AddDefect"
                }, "Insp-BA", this.UserID);
                return RedirectToAction("AddDefect", new { ID = Req.ID });
            }


            // 可以直接執行刪除的圖片
            var DeleteImg = TmpDelete_BAImg.Where(o => o.ImageUKey > 0).ToList();

            if (Req.ListBACriteria != null)
            {
                foreach (BACriteriaItem item in Req.ListBACriteria)
                {
                    var imgs = TmpAdd_BAImg.Where(o => o.BACriteria == item.BACriteria &&
                        !TmpDelete_BAImg.Any(x => x.Seq == o.Seq && x.BACriteria == item.BACriteria)
                    ).ToList();

                    foreach (var t in imgs)
                    {
                        t.Image = t.TempImage;
                    }
                    item.Images = imgs;
                    item.ID = Req.ID;
                }
            }


            InspectionBySP_BA result = _Service.BAProcess(Req, DeleteImg);

            if (!result.ExecuteResult)
            {
                TmpDefectImgSourceList = new List<SelectListItem>();
                TmpAdd_DefectImg = new List<DefectImage>();
                TmpDelete_DefectImg = new List<DefectImage>();

                InspectionBySP_AddDefect model = _Service.GetAddDefectBody(Req.ID);

                model.ErrorMessage = $@"msg.WithError(""{result.ErrorMessage}"")";
                return View(model);

            }
            else
            {
                // 進度更新
                _Service.UpdateSampleRFTInspectionByStep(new SampleRFTInspection()
                {
                    ID = Req.ID,
                    InspectionStep = "Insp-DummyFit"
                }, "Insp-BA", this.UserID);
                return RedirectToAction("DummyFitting", new { ID = Req.ID });
            }

            // 只要Detail UKey底下的圖片有異動，則全部刪除，重新INSERT

        }
        #endregion

        #region DummyFitting 頁面
        public ActionResult DummyFitting(long ID)
        {
            this.CheckSession();


            InspectionBySP_DummyFit model = _Service.GetDummyFit(ID);


            TempData["AllSize"] = model.ArticleSizeList;
            return View(model);
        }

        [HttpPost]
        public ActionResult GetNewSizeByArticle_DF(string Article)
        {
            //TmpMeasurementImgSourceList = new List<SelectListItem>();
            //TmpAdd_MeasurementImg = new List<RFT_Inspection_Measurement_Image>();
            //TmpDelete_MeasurementImg = new List<RFT_Inspection_Measurement_Image>();

            List<ArticleSize> ListSize = (List<ArticleSize>)TempData["AllSize"];

            var listSize = ListSize.Where(o => o.Article == Article).Select(O => O.SizeCode).Distinct().ToList();
            List<SelectListItem> result = new SetListItem().ItemListBinding(listSize);

            // 保存原資料
            TempData["AllSize"] = ListSize;

            return Json(result);
        }

        [HttpPost]
        public ActionResult DummyFitting(InspectionBySP_DummyFit Req, string goPage)
        {
            this.CheckSession();
            InspectionBySP_DummyFit latestModel = new InspectionBySP_DummyFit();
            UpdateModel(latestModel);


            Req.ID = latestModel.ID;
            Req.OrderID = latestModel.OrderID;

            if (goPage == "Back")
            {
                // 進度更新
                _Service.UpdateSampleRFTInspectionByStep(new SampleRFTInspection()
                {
                    ID = Req.ID,
                    InspectionStep = "Insp-BA"
                }, "Insp-DummyFit", this.UserID);
                return RedirectToAction("BeautifulProductAudit", new { ID = Req.ID });
            }
            else
            {
                InspectionBySP_DummyFit model = _Service.DummyFitProcess(Req);

                TempData["AllSize"] = model.ArticleSizeList;

                // 進度更新
                _Service.UpdateSampleRFTInspectionByStep(new SampleRFTInspection()
                {
                    ID = Req.ID,
                    InspectionStep = "Insp-Others"
                }, "Insp-DummyFit", this.UserID);
                return RedirectToAction("Others", new { ID = Req.ID });

            }

        }
        #endregion

        #region Others頁面
        public ActionResult Others(long ID)
        {
            this.CheckSession();

            InspectionBySP_Others model =_Service.Get_Others(ID);
            model.ID = ID;
            model.Inspector = $"{this.UserID} {this.UserName}";
            return View(model);
        }
        [HttpPost]
        public ActionResult Others(InspectionBySP_Others Req, string goPage)
        {
            this.CheckSession();

            InspectionBySP_Others latestModel = new InspectionBySP_Others();
            UpdateModel(latestModel);

            Req.ID = latestModel.ID;

            if (goPage == "Back")
            {
                // 進度更新
                _Service.UpdateSampleRFTInspectionByStep(new SampleRFTInspection()
                {
                    ID = Req.ID,
                    InspectionStep = "Insp-DummyFit"
                }, "Insp-Others", this.UserID);
                return RedirectToAction("DummyFitting", new { ID = Req.ID });
            }

            InspectionBySP_Others result = _Service.OthersProcess(Req);

            if (!result.ExecuteResult)
            {
                InspectionBySP_Others model = _Service.Get_Others(Req.ID);
                model.ID = Req.ID;
                model.Inspector = $"{this.UserID} {this.UserName}";
                return View(model);
            }
            else
            {
                // 進度更新
                _Service.UpdateSampleRFTInspectionByStep(new SampleRFTInspection()
                {
                    ID = Req.ID,
                    InspectionStep = "Submit"
                }, "Insp-Others", this.UserID);
                return RedirectToAction("Index");
            }

        }
        #endregion
    }
}