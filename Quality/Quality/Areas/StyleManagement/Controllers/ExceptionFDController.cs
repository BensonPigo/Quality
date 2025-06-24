using BusinessLogicLayer.Service.StyleManagement;
using DatabaseObject.ViewModel.StyleManagement;
using Quality.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace Quality.Areas.StyleManagement.Controllers
{
    public class ExceptionFDController : BaseController
    {
        public ExceptionFDService Service;

        public ExceptionFDController()
        {
            Service = new ExceptionFDService();
            this.SelectedMenu = "Style Management";
            ViewBag.OnlineHelp = this.OnlineHelp + "StyleManagement.ExceptionFD,,";
        }
        // GET: StyleManagement/ExceptionFD
        public ActionResult Index()
        {
            List<ExceptionFD_ViewModel> model = new List<ExceptionFD_ViewModel>();
            try
            {
                model = Service.GetData();
            }
            catch (Exception ex)
            {
                ViewData["ErrorMessage"] = ex.Message;
            }

            return View(model);
        }

        public ActionResult ToExcel()
        {
            List<ExceptionFD_ViewModel> model = new List<ExceptionFD_ViewModel>();
            try
            {
                model = Service.GetData();
                if (model != null && model.Any())
                {
                    // 1. 在Service層取得資料，生成Excel檔案，放在暫存路徑，回傳檔名

                    string fileMame = Service.GetExcel();

                    // 2. 取得應用程式根目錄下 TMP 的檔案路徑
                    string tempFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                                   "TMP",
                                                   fileMame);

                    // 3. 前端下載方式：請參考Index.cshtml的 「window.location.href = '@download'」;

                    // 載入要下載的檔案
                    byte[] b = System.IO.File.ReadAllBytes(tempFilePath);
                    //清除Response內的HTML
                    Response.Clear();
                    //設定標頭檔資訊 attachment 是本文章的關鍵字
                    Response.AddHeader("Content-Disposition", "attachment;filename=" + fileMame);
                    //開始輸出讀取到的檔案
                    Response.BinaryWrite(b);
                    //一定要加入這一行，否則會持續把Web內的HTML文字也輸出。
                    Response.End();
                    return null;
                }
                else
                {
                    ViewData["ErrorMessage"] = "No Data!";
                    return View("Index", model);
                }
            }
            catch (Exception ex)
            {
                ViewData["ErrorMessage"] = ex.Message;
            }

            return View("Index", model);
        }
    }
}