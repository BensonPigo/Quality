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

    public class BrandGarmentTestController : BaseController
    {
        private BrandGarmentTestService _Service;
        public BrandGarmentTestController()
        {
            _Service = new BrandGarmentTestService();
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.BrandGarmentTest,,";
        }
        // GET: BulkFGT/BrandGarmentTest
        public ActionResult Index()
        {
            BrandGarmentTest_ViewModel model = _Service.GetDefaultModel();

            if (TempData["NewSaveBrandGarmentTestModel"] != null)
            {
                model = (BrandGarmentTest_ViewModel)TempData["NewSaveBrandGarmentTestModel"];
            }
            else if (TempData["EditSaveBrandGarmentTestModel"] != null)
            {
                model = (BrandGarmentTest_ViewModel)TempData["EditSaveBrandGarmentTestModel"];
            }
            else if (TempData["DeleteBrandGarmentTestModel"] != null)
            {
                model = (BrandGarmentTest_ViewModel)TempData["DeleteBrandGarmentTestModel"];
            }

            return View(model);
        }

        /// <summary>
        /// 外部導向至本頁用
        /// </summary>
        public ActionResult IndexGet(string ReportNo, string BrandID, string SeasonID, string StyleID, string Article)
        {
            BrandGarmentTest_ViewModel model = _Service.GetDefaultModel();
            if (string.IsNullOrEmpty(ReportNo) && string.IsNullOrEmpty(ReportNo))
            {
                return View("Index", model);
            }

            BrandGarmentTest_Request request = new BrandGarmentTest_Request()
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
        public ActionResult Query(BrandGarmentTest_Request request)
        {
            BrandGarmentTest_ViewModel model = _Service.GetMainList(request);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            return View("Index", model);
        }

        public ActionResult Detail(BrandGarmentTest_Request request)
        {
            BrandGarmentTest_ViewModel model = _Service.GetDefaultModel(true);

            model = _Service.GetMain(request);

            return View(model);
        }

        public ActionResult New()
        {
            BrandGarmentTest_ViewModel model = _Service.GetDefaultModel(true);

            model.Main.EditType = "New";

            return View("Detail", model);
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "NewSave")]
        public ActionResult NewSave(BrandGarmentTest_ViewModel requestModel)
        {
            BrandGarmentTest_ViewModel model = _Service.NewSave(requestModel, this.MDivisionID, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }
            TempData["NewSaveBrandGarmentTestModel"] = model;

            return RedirectToAction("Detail", new BrandGarmentTest_Request() { ReportNo = model.Main.ReportNo });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "EditSave")]
        public ActionResult EditSave(BrandGarmentTest_ViewModel requestModel)
        {
            BrandGarmentTest_ViewModel model = _Service.EditSave(requestModel, this.UserID);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
            }

            TempData["EditSaveBrandGarmentTestModel"] = model;

            return RedirectToAction("Detail", new BrandGarmentTest_Request() { ReportNo = model.Main.ReportNo });
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        [MultipleButton(Name = "action", Argument = "Delete")]
        public ActionResult Delete(BrandGarmentTest_ViewModel requestModel)
        {
            BrandGarmentTest_ViewModel model = _Service.Delete(requestModel);

            if (!model.Result)
            {
                model.ErrorMessage = $@"msg.WithInfo(""{model.ErrorMessage.Replace("'", string.Empty)}"");";
                return View("Index", model);
            }
            TempData["DeleteBrandGarmentTestModel"] = model;
            return RedirectToAction("Index");
        }


        public ActionResult OrderIDCheck(string orderID)
        {
            BrandGarmentTest_ViewModel model = _Service.GetOrderInfo(orderID);

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
    <input class="""" id=""BrandGarmentTestDoxList_{lastNo}__BrandGarmentTestDoxFile"" name=""BrandGarmentTestDoxList[{lastNo}].BrandGarmentTestDoxFile"" type=""file"" style=""display:none;"">
    <input class="""" id=""BrandGarmentTestDoxList_{lastNo}__IsOldFile"" name=""BrandGarmentTestDoxList[{lastNo}].IsOldFile"" type=""hidden"" value=""false"" readonly=""readonly"">
    <input class="""" id=""BrandGarmentTestDoxList_{lastNo}__ReportNo"" name=""BrandGarmentTestDoxList[{lastNo}].ReportNo"" type=""hidden"" value=""{ReportNo}"" readonly=""readonly"">
    <input class="""" id=""BrandGarmentTestDoxList_{lastNo}__FileName"" name=""BrandGarmentTestDoxList[{lastNo}].FileName"" type=""text"" value=""{FileName}"" readonly=""readonly"">
</div>
<div class=""DetailDataAreaItem1 colBody Row{lastNo}"">
    <input class="""" id=""BrandGarmentTestDoxList_{lastNo}__CreateBy"" name=""BrandGarmentTestDoxList[{lastNo}].CreateBy"" type=""text"" value=""""  readonly=""readonly"">
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
            BrandGarmentTest_ViewModel model = _Service.GetDefaultModel(true);

            model = _Service.GetMain(new BrandGarmentTest_Request() { ReportNo = ReportNo });

            string html = string.Empty;

            foreach (var dox in model.BrandGarmentTestDoxList)
            {
                html += $@"
<div class=""FileListAreaItem1 fileColBody"">
    <input type=""checkbox"" class=""chkBrandGarmentTestDox"" BrandGarmentTestDoxUkey=""{dox.Ukey}"" />
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
        //    BrandGarmentTest_ViewModel model = _Service.GetDefaultModel(true);

        //    model = _Service.Download(new BrandGarmentTest_Request() { ReportNo = ReportNo });

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

        public ActionResult Download(List<BrandGarmentTestDox> brandGarmentTestDoxList)
        {
            BrandGarmentTest_ViewModel model = _Service.GetDefaultModel(true);

            model = _Service.Download(brandGarmentTestDoxList);

            string zipFullName = model.DownloadFileFullName;
            string zipName = model.DownloadFileName;

            string tempFilePath = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + zipName;

            return Json(new { Result = model.Result, ErrMsg = string.Empty, FilaName = tempFilePath });
        }
    
    }
}