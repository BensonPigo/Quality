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

namespace Quality.Areas.BulkFGT.Controllers
{
    public class SearchListController : BaseController
    {
        private ISearchListService _SearchListService;

        public SearchListController()
        {
            _SearchListService = new SearchListService();
            this.SelectedMenu = "Bulk FGT";

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
            if (TempData["Model"] != null)
            {
                model = (SearchList_ViewModel)TempData["Model"];
            }


            return View(model);
        }


        [HttpPost]
        [MultipleButton(Name = "action", Argument = "Query")]
        public ActionResult Query(SearchList_ViewModel Req)
        {
            this.CheckSession();

            // 必填條件
            if ((Req.BrandID == "" || Req.SeasonID == "") || Req.StyleID == "")
            {
                Req.ErrorMessage = $@"
msg.WithInfo('[Style] or [Brand, Season] can't be cmpty. ');
";
                TempData["Model"] = Req;
                return RedirectToAction("Index");
            }


            List<SelectListItem> data = _SearchListService.GetTypeDatasource(this.UserID);

            // Query
            Req.DataList = new List<SearchList_Result>();
            Req = _SearchListService.Get_SearchList(Req);
            Req.TypeDatasource = data;

            if (!Req.Result)
            {
                Req.ErrorMessage = $@"
msg.WithInfo('{Req.ErrorMessage.Replace("'", string.Empty)}');
";
            }

            TempData["Model"] = Req;

            return RedirectToAction("Index");
        }

        [HttpPost]
        public ActionResult ToExcel(SearchList_ViewModel Req)
        {
            this.CheckSession();

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