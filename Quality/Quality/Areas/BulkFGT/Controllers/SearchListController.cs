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
using Quality.Helper;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class SearchListController : BaseController
    {
        private ISearchListService _SearchListService;

        public SearchListController()
        {
            _SearchListService = new SearchListService();
            this.SelectedMenu = "Bulk FGT";
            ViewBag.OnlineHelp = this.OnlineHelp + "BulkFGT.SearchList,,";

        }

        // GET: BulkFGT/SearchList
        public ActionResult Index()
        {
            this.CheckSession();

            List<SelectListItem> data = _SearchListService.GetTypeDatasource(this.UserID);
            SearchList_ViewModel model = new SearchList_ViewModel()
            {
                TypeDatasource = data,
            };
            if (TempData["ModelSearchList"] != null)
            {
                model = (SearchList_ViewModel)TempData["ModelSearchList"];
            }


            return View(model);
        }


        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        [SessionAuthorizeAttribute]
        public ActionResult Query(SearchList_ViewModel Req)
        {
            this.CheckSession();

            // Daily Moisture 的時候，不會篩選Received Date 
            if (!string.IsNullOrEmpty(Req.Type) && Req.Type.Contains("Daily Moisture"))
            {
                Req.ReceivedDate_sText = string.Empty;
                Req.ReceivedDate_eText = string.Empty;
            }
            // 不是Mockup系列 且不是 HT WASH 則清空時間條件
            else if (!string.IsNullOrEmpty(Req.Type) && !Req.Type.ToUpper().Contains("MOCKUP") && !Req.Type.ToUpper().Contains("HT WASH") && !Req.Type.ToUpper().Contains("SALIVA") 
                && !Req.Type.ToUpper().Contains("MOISTURE")
                && !Req.Type.ToUpper().Contains("AGING")
                && !Req.Type.ToUpper().Contains("YELLOWING")
                && !Req.Type.ToUpper().Contains("ABSORBENCY"))

            {
                Req.ReceivedDate_sText = string.Empty;
                Req.ReceivedDate_eText = string.Empty;
                Req.ReportDate_sText = string.Empty;
                Req.ReportDate_eText = string.Empty;
            }

            // 必填條件
            if (string.IsNullOrEmpty(Req.Type))
            {
                Req.ErrorMessage = $@"msg.WithInfo(""[Type] can't be cmpty. "");";
                TempData["ModelSearchList"] = Req;
                return RedirectToAction("Index");
            }

            // 必填條件
            if (!CheckInput(Req))
            {
                Req.ErrorMessage = $@"msg.WithInfo(""Does not include Type, two conditions must be selected. "");";
                TempData["ModelSearchList"] = Req;
                return RedirectToAction("Index");
            }

            List<SelectListItem> data = _SearchListService.GetTypeDatasource(this.UserID);

            Req.MDivisionID = this.MDivisionID;
            // Query
            Req.DataList = new List<SearchList_Result>();
            Req = _SearchListService.Get_SearchList(Req);
            Req.TypeDatasource = data;

            if (!Req.Result)
            {
                Req.ErrorMessage = $@"msg.WithInfo('{(string.IsNullOrEmpty(Req.ErrorMessage) ? string.Empty : Req.ErrorMessage.Replace("'", string.Empty))}');";
            }

            TempData["ModelSearchList"] = Req;

            return RedirectToAction("Index");
        }

        private bool CheckInput(SearchList_ViewModel Req)
        {
            int n = 0;
            if (!string.IsNullOrEmpty(Req.BrandID)) n++;
            if (!string.IsNullOrEmpty(Req.SeasonID)) n++;
            if (!string.IsNullOrEmpty(Req.StyleID)) n++;
            if (!string.IsNullOrEmpty(Req.Article)) n++;
            if (!string.IsNullOrEmpty(Req.Line)) n++;
            if (!string.IsNullOrEmpty(Req.ReceivedDate_sText) && !string.IsNullOrEmpty(Req.ReceivedDate_eText)) n++;
            if (!string.IsNullOrEmpty(Req.ReportDate_sText) && !string.IsNullOrEmpty(Req.ReportDate_eText)) n++;
            if (!string.IsNullOrEmpty(Req.WhseArrival_sText) && !string.IsNullOrEmpty(Req.WhseArrival_eText)) n++;
            return n >= 2;
        }

        [HttpPost]
        [SessionAuthorizeAttribute]
        public ActionResult ToExcel(SearchList_ViewModel Req)
        {
            this.CheckSession();

            Req.MDivisionID = this.MDivisionID;
            SearchList_ViewModel result = _SearchListService.ToExcel(Req);
            result.TempFileName = Request.Url.Scheme + @"://" + Request.Url.Authority + "/TMP/" + result.TempFileName;
            return Json(new { result.Result, result.ErrorMessage, result.TempFileName });
            #region other
            /*
            try
            {

                SearchList_ViewModel model = new SearchList_ViewModel();
                Req.DataList = new List<SearchList_Result>();

                // 取得資料
                model = _SearchListService.Get_SearchList(model);

                XSSFWorkbook book;
                using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "XLT\\Search List.xlsx", FileMode.Open, FileAccess.Read))
                {
                    book = new XSSFWorkbook(file);
                    ISheet sheet = book.GetSheetAt(0);


                    for (int RowIdx = 0; RowIdx < model.DataList.Count; RowIdx++)
                    {
                        //注意，要CreateRoll才對
                        sheet.CreateRow(RowIdx + 1);

                        //注意，要CreateCell才對
                        sheet.GetRow(RowIdx + 1).CreateCell(0).SetCellValue(model.DataList[RowIdx].Type);
                        sheet.GetRow(RowIdx + 1).CreateCell(1).SetCellValue(model.DataList[RowIdx].ReportNo);
                        sheet.GetRow(RowIdx + 1).CreateCell(2).SetCellValue(model.DataList[RowIdx].OrderID);
                        sheet.GetRow(RowIdx + 1).CreateCell(3).SetCellValue(model.DataList[RowIdx].BrandID);
                        sheet.GetRow(RowIdx + 1).CreateCell(4).SetCellValue(model.DataList[RowIdx].StyleID);
                        sheet.GetRow(RowIdx + 1).CreateCell(5).SetCellValue(model.DataList[RowIdx].SeasonID);
                        sheet.GetRow(RowIdx + 1).CreateCell(6).SetCellValue(model.DataList[RowIdx].Article);
                        sheet.GetRow(RowIdx + 1).CreateCell(7).SetCellValue(model.DataList[RowIdx].Artwork);
                        sheet.GetRow(RowIdx + 1).CreateCell(8).SetCellValue(model.DataList[RowIdx].Result);
                        sheet.GetRow(RowIdx + 1).CreateCell(9).SetCellValue(model.DataList[RowIdx].TestDate.Value);
                    }


                }


                using (var ms = new MemoryStream())
                {
                    book.Write(ms);

                    Response.AddHeader("Content-Disposition", $"attachment; filename=CFT Comments_{DateTime.Now.ToString("yyyyMMdd")}.xlsx");
                    Response.BinaryWrite(ms.ToArray());

                    //== 釋放資源
                    book = null;
                    ms.Close();
                    ms.Dispose();

                    Response.Flush();
                    Response.End();
                }



            }
            catch (Exception ex)
            {

                throw ex;
            }
            return View();
            */
            #endregion

        }
    }
}