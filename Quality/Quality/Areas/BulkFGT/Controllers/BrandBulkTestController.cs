using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ViewModel.BulkFGT;
using Microsoft.Office.Interop.Excel;
using Quality.Controllers;
using Quality.Helper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Quality.Helper.Attribute;

namespace Quality.Areas.BulkFGT.Controllers
{

    public class BrandBulkTestController : BaseController
    {
        private BrandBulkTestService _Service;
        public BrandBulkTestController()
        {
            _Service = new BrandBulkTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.BrandBulkTest,,";
        }
        // GET: BulkFGT/BrandBulkTest
        public ActionResult Index()
        {
            BrandBulkTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["NewSaveBrandBulkTestModel"] != null)
            {
                model = (BrandBulkTest_ViewModel)TempData["NewSaveBrandBulkTestModel"];
            }
            else if (TempData["EditSaveBrandBulkTestModel"] != null)
            {
                model = (BrandBulkTest_ViewModel)TempData["EditSaveBrandBulkTestModel"];
            }
            else if (TempData["DeleteBrandBulkTestModel"] != null)
            {
                model = (BrandBulkTest_ViewModel)TempData["DeleteBrandBulkTestModel"];
            }

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            BrandBulkTest_ViewModel model = _Service.GetDefaultModel();
            if (string.IsNullOrEmpty(ReportNo) && string.IsNullOrEmpty(ReportNo))
            {
                return View("Index", model);
            }

            BrandBulkTest_Request request = new BrandBulkTest_Request()
            {
                ReportNo = ReportNo,
                BrandID = BrandID,
                SeasonID = SeasonID,
                StyleID = StyleID,
                Article = Article,
            };

            model = _Service.GetMainList(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }
        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(BrandBulkTest_Request request)
        {
            BrandBulkTest_ViewModel model = _Service.GetMainList(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult Detail(BrandBulkTest_Request request)
        {
            BrandBulkTest_ViewModel model = _Service.GetDefaultModel(true);

            model = _Service.GetMain(request);

            return View(model);
        }

        public ActionResult New()
        {
            BrandBulkTest_ViewModel model = _Service.GetDefaultModel(true);

            model.Main.EditType = "New";

            return View("Detail", model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(BrandBulkTest_ViewModel requestModel)
        {
            BrandBulkTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSaveBrandBulkTestModel"] = model;

            return RedirectToAction("Detail", new BrandBulkTest_Request() { ReportNo = model.Main.ReportNo });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(BrandBulkTest_ViewModel requestModel)
        {
            BrandBulkTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSaveBrandBulkTestModel"] = model;

            return RedirectToAction("Detail", new BrandBulkTest_Request() { ReportNo = model.Main.ReportNo });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(BrandBulkTest_ViewModel requestModel)
        {
            BrandBulkTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }
            TempData["DeleteBrandBulkTestModel"] = model;
            return RedirectToAction("Index");
        }


        public ActionResult OrderIDCheck(string orderID)
        {
            BrandBulkTest_ViewModel model = _Service.GetOrderInfo(orderID);

            if (!model.Result)
            {
                return Json(new { ErrMsg = model.ErrorMessage, Result = model.Result });
            }
            else
            {
                List<string> Article_Source = new List<string>();
                foreach (var item in model.Article_Source)
                {
                    string selected = string.Empty;
                    if (item.Value == model.Main.Article)
                    {
                        selected = "selected";
                    }
                    Article_Source.Add($"<option {selected} value='{item.Value}'>{item.Text}</option>");
                }

                return Json(new { ErrMsg = string.Empty, Result = model.Result, Main = model.Main, ArticleSource = Article_Source });
            }
        }


        public ActionResult AddFileRow(int lastNo, string FileName, string ReportNo)
        {

            string html = string.Empty;
            html += $@"
<!--#region Row {lastNo}-->

<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input type=""hidden"" class=""detailRowIdx"" name=""name"" value=""{lastNo}"" readonly=""readonly"">
    <input class="""" id=""BrandBulkTestDoxList_{lastNo}__BrandBulkTestDoxFile"" name=""BrandBulkTestDoxList[{lastNo}].BrandBulkTestDoxFile"" type=""file"" style=""display:none;"">
    <input class="""" id=""BrandBulkTestDoxList_{lastNo}__IsOldFile"" name=""BrandBulkTestDoxList[{lastNo}].IsOldFile"" type=""hidden"" value=""false"" readonly=""readonly"">
    <input class="""" id=""BrandBulkTestDoxList_{lastNo}__ReportNo"" name=""BrandBulkTestDoxList[{lastNo}].ReportNo"" type=""hidden"" value=""{ReportNo}"" readonly=""readonly"">
    <input class="""" id=""BrandBulkTestDoxList_{lastNo}__FileName"" name=""BrandBulkTestDoxList[{lastNo}].FileName"" type=""text"" value=""{FileName}"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""BrandBulkTestDoxList_{lastNo}__CreateBy"" name=""BrandBulkTestDoxList[{lastNo}].CreateBy"" type=""text"" value=""""  readonly=""readonly"">
</div>

<div class=""DetailDataAreaItem2 colBody Row{lastNo}"">
    <img class='detailDelete' src=""/Image/Icon/Delete.png"" width=""30"">
</div>

<!--#endregion-->
";


            return Content(html);
        }

        public ActionResult GetIndexFileRow(string ReportNo)
        {
            BrandBulkTest_ViewModel model = _Service.GetDefaultModel(true);

            model = _Service.GetMain(new BrandBulkTest_Request() { ReportNo = ReportNo });

            string html = string.Empty;

            foreach (var dox in model.BrandBulkTestDoxList)
            {
                html += $@"
<div class=""FileListAreaItem1 fileColBody"">
    <input type=""checkbox"" class=""chkBrandBulkTestDox"" BrandBulkTestDoxUkey=""{dox.Ukey}"" />
</div>

<div class=""FileListAreaItem2 fileColBody"">
    {dox.FileName}
</div>

";
            }

            return Content(html);
        }
        //public ActionResult Download(string ReportNo)
        //{
        //    BrandBulkTest_ViewModel model = _Service.GetDefaultModel(true);

        //    model = _Service.Download(new BrandBulkTest_Request() { ReportNo = ReportNo });

        //    string zipFullName = model.DownloadFileFullName;
        //    string zipName = model.DownloadFileName;
        //    var tempFile = System.IO.Path.Combine(zipFullName);

        //    using (MemoryStream obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(zipFullName)))
        //    {
        //        Response.ContentType = "application/zip";
        //        Response.AddHeader("Content-Disposition", $"attachment; filename={zipName}");
        //        Response.BinaryWrite(obj_stream.ToArray());
        //        obj_stream.Close();
        //        obj_stream.Dispose();
        //        Response.Flush();
        //        Response.End();
        //    }

        //    return null;
        //}

        public ActionResult Download(List<BrandBulkTestDox> brandBulkTestDoxList)
        {
            BrandBulkTest_ViewModel model = _Service.GetDefaultModel(true);

            model = _Service.Download(brandBulkTestDoxList);

            string zipFullName = model.DownloadFileFullName;
            string zipName = model.DownloadFileName;

            string tempFilePath = "/TMP/" + zipName;

            return Json(new { Result = model.Result, ErrMsg = string.Empty, FilaName = tempFilePath });
        }
    
    }
}