using BusinessLogicLayer.Service.StyleManagement;
using DatabaseObject;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.SampleRFT;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.StyleManagement.Controllers
{
    public class StyleResultController : BaseController
    {
        private StyleResultService _Service;

        public StyleResultController()
        {
            _Service = new StyleResultService();
            this.SelectedMenu = "Style Management";
            ViewBag.OnlineHelp = this.OnlineHelp + "StyleManagement.StyleResult,,";
        }
        // GET: StyleManagement/StyleResult
        public ActionResult Index()
        {
            StyleResult_ViewModel model = new StyleResult_ViewModel()
            {
                SampleRFT = new List<StyleResult_SampleRFT>(),
                FTYDisclaimer = new List<StyleResult_FTYDisclaimer>() { new StyleResult_FTYDisclaimer() },
                RRLR = new List<StyleResult_RRLR>() { new StyleResult_RRLR() },
                BulkFGT = new List<StyleResult_BulkFGT>() { new StyleResult_BulkFGT() }
            };

            if (TempData["ModelStyleResult"] != null)
            {
                model = (StyleResult_ViewModel)TempData["ModelStyleResult"];
            }
            if (TempData["tempFilePath"] != null)
            {
                ViewData["tempFilePath"] = TempData["tempFilePath"].ToString();
            }

            return View(model);
        }

        public ActionResult BarcodeScan()
        {
            this.CheckSession();

            return View();
        }


        public ActionResult IndexGet(string BrandID, string SeasonID, string StyleID)
        {
            this.CheckSession();
            StyleResult_ViewModel model = new StyleResult_ViewModel();
            StyleResult_Request Req = new StyleResult_Request() 
            {
                BrandID= BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
            };

            if (Req == null || string.IsNullOrEmpty(Req.StyleID) || string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID))
            {
                model.MsgScript = $@"msg.WithInfo(""Style, Brand and Season cannot be empty."");";
                return View("Index", model);
            }

            Req.MDivisionID = this.MDivisionID;
            model = _Service.Get_StyleResult_Browse(Req);
            return View("Index", model);
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(StyleResult_Request Req)
        {
            this.CheckSession();
            StyleResult_ViewModel model = new StyleResult_ViewModel() 
            {
                SampleRFT = new List<StyleResult_SampleRFT>(),
                FTYDisclaimer = new List<StyleResult_FTYDisclaimer>(),
                RRLR = new List<StyleResult_RRLR>(),
                BulkFGT = new List<StyleResult_BulkFGT>(),
            };
            if (Req == null || string.IsNullOrEmpty(Req.StyleID) || string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID))
            {
                model.MsgScript = $@"msg.WithInfo(""Style, Brand and Season cannot be empty."");";
                return View("Index", model);
            }

            Req.MDivisionID = this.MDivisionID;
            model = _Service.Get_StyleResult_Browse(Req);

            model.SampleRFT = model.SampleRFT == null ? new List<StyleResult_SampleRFT>() : model.SampleRFT;
            model.FTYDisclaimer = model.FTYDisclaimer == null ? new List<StyleResult_FTYDisclaimer>() : model.FTYDisclaimer;
            model.RRLR = model.RRLR == null ? new List<StyleResult_RRLR>() : model.RRLR;
            model.BulkFGT = model.BulkFGT == null ? new List<StyleResult_BulkFGT>() : model.BulkFGT;

            return View("Index", model);
        }

        [HttpPost]
        public ActionResult CheckStyle(StyleResult_Request Req)
        {
            this.CheckSession();

            try
            {
                var list = _Service.GetStyle(new StyleResult_Request() { StyleUkey = Req.StyleUkey }).ToList();

                if (list.Any())
                {
                    return Json(list.FirstOrDefault());
                }
                else
                {
                    return Json(new List<StyleResult_Request>());
                }
            }
            catch (Exception ex)
            {
                return Json(ex.Message);
            }
        }

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "SampleRFTToExcel")]
        public ActionResult SampleRFT_ToExcel(StyleResult_Request Req)
        {
            this.CheckSession();

            StyleResult_ViewModel model = new StyleResult_ViewModel();

            // 1. 在Service層取得資料，生成Excel檔案，放在暫存路徑，回傳檔名
            model = _Service.GetExcel(Req);
            string tempFilePath = model.TempFileName;

            // 2. 取得hotst name，串成下載URL ，傳到準備前端下載
            tempFilePath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + tempFilePath;

            // 3. 前端下載方式：請參考Index.cshtml的 「window.location.href = '@download'」;            

            TempData["tempFilePath"] = tempFilePath;
            TempData["ModelStyleResult"] = model;

            return RedirectToAction("Index");
        }

        public ActionResult DownloadFDFile(string FDFilePath, string FileName)
        {

            //宣告並建立WebClient物件
            WebClient wc = new WebClient();
            //載入要下載的檔案
            byte[] b = wc.DownloadData(FDFilePath);
            //清除Response內的HTML
            Response.Clear();
            //設定標頭檔資訊 attachment 是本文章的關鍵字
            Response.AddHeader("Content-Disposition", "attachment;filename=" + FileName);
            //開始輸出讀取到的檔案
            Response.BinaryWrite(b);
            //一定要加入這一行，否則會持續把Web內的HTML文字也輸出。
            Response.End();

            return null;

        }


        public ActionResult DownloadRRLRFile(string FilePath,string BrandID, string SeasonID, string StyleID)
        {

            string FileName = $"{BrandID}_{SeasonID}_{StyleID}.xlsx";
            //宣告並建立WebClient物件
            WebClient wc = new WebClient();


            if (System.IO.File.Exists(FilePath + FileName))
            {
                //載入要下載的檔案
                byte[] b = wc.DownloadData(FilePath + FileName);
                //清除Response內的HTML
                Response.Clear();
                //設定標頭檔資訊 attachment 是本文章的關鍵字
                Response.AddHeader("Content-Disposition", "attachment;filename=" + FileName);
                //開始輸出讀取到的檔案
                Response.BinaryWrite(b);
                //一定要加入這一行，否則會持續把Web內的HTML文字也輸出。
                Response.End();
                return null;
            }
            else
            {
                StyleResult_ViewModel model = new StyleResult_ViewModel();
                StyleResult_Request Req = new StyleResult_Request()
                {
                    BrandID = BrandID,
                    SeasonID = SeasonID,
                    StyleID = StyleID,
                };

                Req.MDivisionID = this.MDivisionID;
                model = _Service.Get_StyleResult_Browse(Req);
                model.MsgScript = $@"msg.WithInfo(""Could not found RR/LR Report."");";

                return View("Index", model);
            }



        }
    }
}