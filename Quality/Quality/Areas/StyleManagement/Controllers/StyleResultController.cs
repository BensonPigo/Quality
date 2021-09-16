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
        }
        // GET: StyleManagement/StyleResult
        public ActionResult Index()
        {
            StyleResult_ViewModel model = new StyleResult_ViewModel()
            {
                SampleRFT = new List<StyleResult_SampleRFT>(),
                FTYDisclamier = new List<StyleResult_FTYDisclamier>() { new StyleResult_FTYDisclamier() },
                RRLR = new List<StyleResult_RRLR>() { new StyleResult_RRLR() },
                BulkFGT = new List<StyleResult_BulkFGT>() { new StyleResult_BulkFGT() }
            };

            if (TempData["Model"] != null)
            {
                model = (StyleResult_ViewModel)TempData["Model"];
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

        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(StyleResult_Request Req)
        {
            this.CheckSession();
            StyleResult_ViewModel model = new StyleResult_ViewModel();
            if (Req == null || string.IsNullOrEmpty(Req.StyleID) || string.IsNullOrEmpty(Req.BrandID) || string.IsNullOrEmpty(Req.SeasonID))
            {
                model.MsgScript = $@"
msg.WithInfo('Style, Brand and Season cannot be empty.');
";
                return View("Index", model);
            }

            Req.MDivisionID = this.MDivisionID;
            model = _Service.Get_StyleResult_Browse(Req);
            return View("Index", model);
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
            TempData["Model"] = model;

            return RedirectToAction("Index");
        }

        public ActionResult DownLoadFDFile(string FDFilePath, string FileName)
        {


            //設定要下載的檔案路徑 及 儲存的檔名
            string path = "c:\\temp\\123.pdf";
            //string FileName = "abc.pdf";
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
            //string fileName = FDFile;

            //MemoryStream obj_stream = new MemoryStream();
            //var tempFile = FDFile;
            //obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(tempFile));
            //Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
            //Response.BinaryWrite(obj_stream.ToArray());
            //obj_stream.Close();
            //obj_stream.Dispose();
            //Response.Flush();
            //Response.End();
            //return null;

        }
    }
}