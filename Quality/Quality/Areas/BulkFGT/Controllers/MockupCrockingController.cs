using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using Quality.Controllers;
using System;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class MockupCrockingController : BaseController
    {
        private IMockupCrockingService _MockupCrockingService;

        // GET: BulkFGT/MockupCrocking
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 使用Microsoft.Office.Interop.Excel的寫法
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        [HttpPost]
        [MultipleButton(Name = "action", Argument = "ToPDF")]
        public ActionResult ToPDF(MockupCrocking_ViewModel Req)
        {
            this.CheckSession();
            try
            {
                MockupCrocking_ViewModel model = new MockupCrocking_ViewModel();
                if (model.ReportNo == null || string.IsNullOrEmpty(model.ReportNo))
                {
                    Req.ErrorMessage = $@"
msg.WithInfo('ReportNo cannot be empty!');
";
                    return View("Index", Req);
                }

                model = _MockupCrockingService.GetMockupCrocking(new MockupCrocking() { ReportNo = model.ReportNo });
                if (model.MockupCrocking == null || model.MockupCrocking.Count == 0)
                {
                    Req.ErrorMessage = $@"
msg.WithInfo('Cannot found ReportNo {Req.ReportNo}');
";
                    return View("Index", Req);
                }

                // 1. 在Service層取得資料，生成Excel檔案，放在暫存路徑，回傳檔名

                MockupCrocking_ViewModel result = _MockupCrockingService.GetExcel(model);
                string tempFilePath = result.TempFileName;
                // 2. 取得hotst name，串成下載URL ，傳到準備前端下載
                // URL範例：https://misap:1880/TMP/CFT Comments20210826f7f4ad14-186f-451a-9bc1-6edbcaf6cd65.xlsx 
                // (暫存檔檔名是CFT Comments20210826f7f4ad14-186f-451a-9bc1-6edbcaf6cd65.xlsx)
                tempFilePath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + tempFilePath;

                // 3. 前端下載方式：請參考Index.cshtml的 「window.location.href = '@download'」;

                TempData["tempFilePath"] = tempFilePath;
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return RedirectToAction("Index");
        }
    }
}