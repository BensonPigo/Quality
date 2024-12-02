using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Org.BouncyCastle.Ocsp;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

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
                string basefileName = "Search List";
                string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xlsx";
                Microsoft.Office.Interop.Excel.Application objApp = MyUtility.Excel.ConnectExcel(openfilepath);
                objApp.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = objApp.ActiveWorkbook.Worksheets[1]; // 取得工作表

                for (int i = 0; i <= model.DataList.Count - 1; i++)
                {
                    worksheet.Cells[i + 2, 1] = model.DataList[i].Type;
                    worksheet.Cells[i + 2, 2] = model.DataList[i].ReportNo;
                    worksheet.Cells[i + 2, 3] = model.DataList[i].OrderID;
                    worksheet.Cells[i + 2, 4] = model.DataList[i].BrandID;
                    worksheet.Cells[i + 2, 5] = model.DataList[i].StyleID;
                    worksheet.Cells[i + 2, 6] = model.DataList[i].SeasonID;
                    worksheet.Cells[i + 2, 7] = model.DataList[i].Article;
                    worksheet.Cells[i + 2, 8] = model.DataList[i].Line;
                    worksheet.Cells[i + 2, 9] = model.DataList[i].Artwork;
                    worksheet.Cells[i + 2, 10] = model.DataList[i].Result;
                    worksheet.Cells[i + 2, 11] = model.DataList[i].ReceivedDate.HasValue ? model.DataList[i].ReceivedDate.Value.ToString("yyyy/MM/dd HH:mm:ss") : string.Empty;
                    worksheet.Cells[i + 2, 12] = model.DataList[i].ReportDate.HasValue ? model.DataList[i].ReportDate.Value.ToString("yyyy/MM/dd HH:mm:ss") : string.Empty;
                    worksheet.Cells[i + 2, 13] = model.DataList[i].TestDate.HasValue ? model.DataList[i].TestDate.Value.ToString("yyyy/MM/dd HH:mm:ss") : string.Empty;
                    worksheet.Cells[i + 2, 14] = model.DataList[i].AddName;
                }

                worksheet.Columns.AutoFit();

                // Save Excel
                string fileName = $"{basefileName}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath;
                filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                Excel.Workbook workbook = objApp.ActiveWorkbook;
                workbook.SaveAs(filepath);
                workbook.Close();
                objApp.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(objApp);

                model.Result = true;
                model.TempFileName = fileName;
            }
            else
            {
                model.Result = false;
                model.TempFileName = "";
            }

            return model;
        }

    }
}
