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
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class PullingTestController : BaseController
    {
        public PullingTestService Service;

        public PullingTestController()
        {
            Service = new PullingTestService();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.PullingTest,,";
        }

        // GET: BulkFGT/PullingTest
        public ActionResult Index()
        {
            PullingTest_ViewModel model = new PullingTest_ViewModel()
            {
                BrandID = string.Empty,
                SeasonID = string.Empty,
                StyleID = string.Empty,
                Article = string.Empty,
                ReportNo_Source = new List<SelectListItem>(),
                Detail = new PullingTest_Result(),
            };

            ViewBag.FactoryID = this.FactoryID;
            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            PullingTest_ViewModel Req = new PullingTest_ViewModel()
            {
                ReportNo_Query= ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            PullingTest_ViewModel model = Service.GetReportNoList(Req);

            if (model.ReportNo_Source != null && model.ReportNo_Source.Any())
            {
                if (string.IsNullOrEmpty(model.ReportNo_Query))
                {
                    model.ReportNo_Query = model.ReportNo_Source.FirstOrDefault().Value;
                }

                model.Detail = Service.GetData(model.ReportNo_Query).Detail;
            }

            ViewBag.FactoryID = this.FactoryID;
            return View("Index", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(PullingTest_ViewModel Req)
        {
            PullingTest_ViewModel model = Service.GetReportNoList(Req);

            if (model.ReportNo_Source != null && model.ReportNo_Source.Any())
            {
                if (!string.IsNullOrEmpty(model.ReportNo_Query))
                {
                    model.ReportNo_Query = model.ReportNo_Source.FirstOrDefault().Value;
                }

                model.Detail = Service.GetData(model.ReportNo_Query).Detail;
            }

            ViewBag.FactoryID = this.FactoryID;
            return View("Index", model);
        }

        [HttpPost]
        public ActionResult CheckSP(string POID)
        {
            PullingTest_Result o = new PullingTest_Result();
            try
            {
                o = Service.CheckSP(POID);

            }
            catch (Exception ex)
            {
                return Json(ex);
            }

            return Json(o);
        }

        [HttpPost]
        public ActionResult GetPullUnit(string BrandID)
        {
            PullingTest_Result o = new PullingTest_Result();
            try
            {
                o = Service.GetPullUnit(BrandID);

            }
            catch (Exception ex)
            {
                return Json(ex);
            }

            return Json(o);
        }

        [HttpPost]
        public ActionResult GetStandard(string BrandID, string TestItem, string PullForceUnit)
        {
            PullingTest_Result o = new PullingTest_Result();
            try
            {
                o = Service.GetStandard(BrandID, TestItem, PullForceUnit);

            }
            catch (Exception ex)
            {
                return Json(ex);
            }

            return Json(o);
        }

        [HttpPost]
        public ActionResult GetDetail(string ReportNo)
        {
            PullingTest_ViewModel model = new PullingTest_ViewModel();
            try
            {
                model = Service.GetData(ReportNo);


                var BeforeImage = model.Detail.TestBeforePicture is null ? new byte[1] : model.Detail.TestBeforePicture;
                var base64_Before = Convert.ToBase64String(BeforeImage);
                var imgSrc_Before = String.Format("data:image/png;base64,{0}", base64_Before);

                var AfterImage = model.Detail.TestAfterPicture is null ? new byte[1] : model.Detail.TestAfterPicture;
                var base64_After = Convert.ToBase64String(AfterImage);
                var imgSrc_After = String.Format("data:image/png;base64,{0}", base64_After);

                model.Detail.TestBeforePicture_Base64 = imgSrc_Before;
                model.Detail.TestAfterPicture_Base64 = imgSrc_After;

                model.Result = true;

            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;

            }
            return Json(model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Edit")]
        public ActionResult EditSave(PullingTest_ViewModel Req)
        {
            Req.Detail.EditName = this.UserID;

            bool IsSendMail = Req.Detail.Result == "Fail";
            string ToAddress = Req.ToAddress;
            string CcAddress = Req.CcAddress;
            string PullForceUnit = Req.Detail.PullForceUnit;

            //準備回傳的model
            PullingTest_ViewModel model = new PullingTest_ViewModel()
            {
                Detail = new PullingTest_Result(),
                TestItem_Source = new List<SelectListItem>(),
            };

            // 更新
            model = Service.Update(Req.Detail);

            if (model.Result)
            {
                //model = Service.GetReportNoList(Req);
                //存檔後重新取得ReportNo清單
                model.Detail = Service.GetData(Req.Detail.ReportNo).Detail;

                model.ReportNo_Query = Req.Detail.ReportNo;

                model.BrandID = model.Detail.BrandID;
                model.SeasonID = model.Detail.SeasonID;
                model.StyleID = model.Detail.StyleID;
                model.Article = model.Detail.Article;

                model = Service.GetReportNoList(model);
            }
            
            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.ToString().Replace("\r\n", "<br />")}"");";
            }
            else if (model.Detail.Result == "Fail")
            {
                model.ErrorMessage = "FailMail();";
            }


            ViewBag.FactoryID = this.FactoryID;
            return View("Index", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "New")]
        public ActionResult NewSave(PullingTest_ViewModel Req)
        {
            bool IsSendMail = Req.Detail.Result == "Fail";
            string ToAddress = Req.ToAddress;
            string CcAddress = Req.CcAddress;
            string PullForceUnit = Req.Detail.PullForceUnit;

            //M (3 碼) + PT + 年 (2 碼) + 月 (2 碼) + 流水號 (4 碼)
            Req.Detail.ReportNo = this.MDivisionID + "PT" + DateTime.Now.ToString("yyyyyMM").Substring(3, 4);
            Req.Detail.AddName = this.UserID;

            //準備回傳的model
            PullingTest_ViewModel model = new PullingTest_ViewModel()
            {
                BrandID = Req.Detail.BrandID,
                SeasonID = Req.Detail.SeasonID,
                StyleID = Req.Detail.StyleID,
                Article = Req.Detail.Article,
                Detail = new PullingTest_Result(),
                TestItem_Source = new List<SelectListItem>(),
            };

            // 新增
            model = Service.Insert(Req.Detail);

            if (model.Result)
            {

                // 新增成功後，搜尋條件帶入：這次新增的BrandID、SeasonID、StyleID、Article，並帶出ReportNo下拉選單
                PullingTest_ViewModel newQueryModel = new PullingTest_ViewModel()
                {
                    BrandID = Req.Detail.BrandID,
                    SeasonID = Req.Detail.SeasonID,
                    StyleID = Req.Detail.StyleID,
                    Article = Req.Detail.Article,
                };

                //存檔後搜尋結果
                model = Service.GetReportNoList(newQueryModel);
                model.ReportNo_Source.FirstOrDefault().Selected = true;

                model.Detail = Service.GetData(model.ReportNo_Source.FirstOrDefault().Value).Detail;
            }


            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.ToString().Replace("\r\n", "<br />")}"");";
            }
            else if (model.Detail.Result == "Fail")
            {
                model.ErrorMessage = "FailMail();";
            }

            model.ReportNo_Query = model.Detail.ReportNo;

            ViewBag.FactoryID = this.FactoryID;
            return View("Index", model);

        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(PullingTest_ViewModel Req)
        {
            string ReportNo = Req.Detail.ReportNo;

            //準備回傳的model
            PullingTest_ViewModel Result = new PullingTest_ViewModel()
            {
                BrandID = Req.BrandID,
                SeasonID = Req.SeasonID,
                StyleID = Req.StyleID,
                Article = Req.Article,
                TestItem_Source = new List<SelectListItem>(),
            };

            // 刪除
            Result = Service.Delete(ReportNo);

            if (Result.Result)
            {
                //刪除後重新取得ReportNo清單
                Result = Service.GetReportNoList(Req);
                Result.ReportNo_Query = "";
                Result.Detail = new PullingTest_Result();
                Result.ReportNo_Query = "";

                if (Result.ReportNo_Source != null && Result.ReportNo_Source.Any())
                {
                    if (string.IsNullOrEmpty(Result.ReportNo_Query))
                    {
                        Result.ReportNo_Query = Result.ReportNo_Source.FirstOrDefault().Value;
                    }

                    Result.Detail = Service.GetData(Result.ReportNo_Query).Detail;
                }

            }

            ViewBag.FactoryID = this.FactoryID;
            return View("Index", Result);
        }

        [HttpPost]
        public JsonResult FailMail(string ReportNo, string TO, string CC)
        {
            SendMail_Result result = Service.FailSendMail(ReportNo, TO, CC);
            return Json(result);
        }
    }
}