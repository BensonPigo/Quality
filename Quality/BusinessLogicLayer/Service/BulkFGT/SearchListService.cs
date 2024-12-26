using BusinessLogicLayer.Interface.BulkFGT;
using ClosedXML.Excel;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class SearchListService : ISearchListService
    {
        private ISearchListProvider SearchListProvider;

        public List<SelectListItem> GetTypeDatasource(string Pass1ID, bool check)
        {
            List<SelectListItem> result = new List<SelectListItem>();
            try
            {
                SearchListProvider = new SearchListProvider(Common.ManufacturingExecutionDataAccessLayer);
                result = SearchListProvider.GetTypeDatasource(Pass1ID, check).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }

        public SearchList_ViewModel Get_SearchList(SearchList_ViewModel Req)
        {
            Req.DataList = new List<SearchList_Result>();
            SearchList_ViewModel model = Req;

            try
            {
                SearchListProvider = new SearchListProvider(Common.ProductionDataAccessLayer);
                var res = SearchListProvider.Get_SearchList(Req).ToList();

                model.DataList = res;
                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public SearchList_ViewModel ToExcel(SearchList_ViewModel Req)
        {
            SearchListProvider = new SearchListProvider(Common.ProductionDataAccessLayer);
            SearchList_ViewModel model = Get_SearchList(Req);

            if (model.DataList.Count > 0)
            {
                string baseFileName = "Search List";
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string fileName = $"{baseFileName}_{DateTime.Now:yyyyMMdd}_{Guid.NewGuid()}.xlsx";
                string filePath = Path.Combine(baseFilePath, "TMP", fileName);

                string templatePath = Path.Combine(baseFilePath, "XLT", "Search List.xlsx");
                // 建立 ClosedXML 工作簿
                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    // 建立標題列
                    worksheet.Cell(1, 1).Value = "Type";
                    worksheet.Cell(1, 2).Value = "ReportNo";
                    worksheet.Cell(1, 3).Value = "OrderID";
                    worksheet.Cell(1, 4).Value = "BrandID";
                    worksheet.Cell(1, 5).Value = "StyleID";
                    worksheet.Cell(1, 6).Value = "SeasonID";
                    worksheet.Cell(1, 7).Value = "Article";
                    worksheet.Cell(1, 8).Value = "Line";
                    worksheet.Cell(1, 9).Value = "Artwork";
                    worksheet.Cell(1, 10).Value = "Result";
                    worksheet.Cell(1, 11).Value = "ReceivedDate";
                    worksheet.Cell(1, 12).Value = "ReportDate";
                    worksheet.Cell(1, 13).Value = "TestDate";
                    worksheet.Cell(1, 14).Value = "AddName";

                    // 填入資料列
                    for (int i = 0; i < model.DataList.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = model.DataList[i].Type;
                        worksheet.Cell(i + 2, 2).Value = model.DataList[i].ReportNo;
                        worksheet.Cell(i + 2, 3).Value = model.DataList[i].OrderID;
                        worksheet.Cell(i + 2, 4).Value = model.DataList[i].BrandID;
                        worksheet.Cell(i + 2, 5).Value = model.DataList[i].StyleID;
                        worksheet.Cell(i + 2, 6).Value = model.DataList[i].SeasonID;
                        worksheet.Cell(i + 2, 7).Value = model.DataList[i].Article;
                        worksheet.Cell(i + 2, 8).Value = model.DataList[i].Line;
                        worksheet.Cell(i + 2, 9).Value = model.DataList[i].Artwork;
                        worksheet.Cell(i + 2, 10).Value = model.DataList[i].Result;
                        worksheet.Cell(i + 2, 11).Value = model.DataList[i].ReceivedDate?.ToString("yyyy/MM/dd HH:mm:ss");
                        worksheet.Cell(i + 2, 12).Value = model.DataList[i].ReportDate?.ToString("yyyy/MM/dd HH:mm:ss");
                        worksheet.Cell(i + 2, 13).Value = model.DataList[i].TestDate?.ToString("yyyy/MM/dd HH:mm:ss");
                        worksheet.Cell(i + 2, 14).Value = model.DataList[i].AddName;
                    }

                    // 自動調整欄寬
                    worksheet.Columns().AdjustToContents();

                    // 儲存 Excel 檔案
                    workbook.SaveAs(filePath);
                }

                model.Result = true;
                model.TempFileName = fileName;
            }
            else
            {
                model.Result = false;
                model.TempFileName = string.Empty;
            }

            return model;
        }

        private void AddImageToWorksheet(IXLWorksheet worksheet, byte[] imageData, int row, int col, int width, int height)
        {
            if (imageData != null)
            {
                using (var stream = new MemoryStream(imageData))
                {
                    worksheet.AddPicture(stream)
                             .MoveTo(worksheet.Cell(row, col), 5, 5)
                             .WithSize(width, height);
                }
            }
        }

    }
}
