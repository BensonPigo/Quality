using DatabaseObject.ViewModel;
using System.Collections.Generic;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System.Linq;
using BusinessLogicLayer.Interface.SampleRFT;
using System.IO;
using Ict;
using Sci.Data;
using System;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Sci;

namespace BusinessLogicLayer.Service.SampleRFT
{
    public class CFTCommentsService : ICFTCommentsService
    {
        private ICFTCommentsProvider _CFTCommentsProvider;

        public CFTComments_ViewModel Get_CFT_Orders(CFTComments_ViewModel Req)
        {
            CFTComments_ViewModel model = new CFTComments_ViewModel();
            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ProductionDataAccessLayer);
                var res = _CFTCommentsProvider.Get_CFT_Orders(Req).ToList();

                if (res.Any())
                {
                    if (Req.QueryType == "Style")
                    {
                        model.BrandID = Req.BrandID;
                        model.SeasonID = Req.SeasonID;
                        model.StyleID = Req.StyleID;
                    }

                    if (Req.QueryType == "OrderID")
                    {
                        model = res.FirstOrDefault();
                    }
                }

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public CFTComments_ViewModel Get_CFT_OrderComments(CFTComments_ViewModel Req)
        {
            Req.DataList = new List<CFTComments_Result>();
            CFTComments_ViewModel model = Req;

            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);
                var res = _CFTCommentsProvider.Get_CFT_OrderComments(Req).ToList();

                foreach (var item in res)
                {
                    item.Comnments = item.Comnments == null ? string.Empty : item.Comnments.Replace("*", "<br>");
                }

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

        /// <summary>
        /// 廢棄：取得Datatable
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public DataTable GetExcel_DataTable(CFTComments_ViewModel Req)
        {
            string TempTilePath = string.Empty;
            CFTComments_ViewModel result = new CFTComments_ViewModel();
            DataTable dt = new DataTable();
            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得Datatable
                dt = _CFTCommentsProvider.Get_CFT_OrderComments_DataTable(Req);

            }
            catch (Exception ex)
            {
                throw ex;
                //result.Result = false;
                //result.ErrorMessage = ex.Message;
            }

            return dt;
        }

        /// <summary>
        /// 透過  Microsoft.Office.Interop.Excel 生成Excel，並下載至暫存路徑
        /// </summary>
        /// <param name="Req"></param>
        /// <returns>暫存檔路徑</returns>
        public CFTComments_ViewModel GetExcel(CFTComments_ViewModel Req)
        {
            string TempTilePath = string.Empty;
            CFTComments_ViewModel result = new CFTComments_ViewModel();
            DataTable dt = new DataTable();
            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);
                string[] aryAlpha = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM" };

                // 取得Datatable
                dt = _CFTCommentsProvider.Get_CFT_OrderComments_DataTable(Req);

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                }

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                }

                if (dt.Rows.Count == 0) 
                {
                    result.TempFileName = string.Empty;
                    result.Result = false;
                    return result;
                }

                // 開啟excel app
                Excel.Application excelApp = MyUtility.Excel.ConnectExcel(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\CFT Comments.xltx");

                Excel.Worksheet worksheet = excelApp.Sheets[1];

                int RowIdx = 0;
                List<string> sampleStages = dt.AsEnumerable().Select(x => x.Field<string>("SampleStage")).OrderBy(x => x).Distinct().ToList();
                foreach (string sampleStage in sampleStages)
                {
                    worksheet.Cells[1, 2 + RowIdx] = sampleStage;
                    RowIdx++;
                }

                worksheet.Range["B1", $"{aryAlpha[sampleStages.Count]}1"].Interior.Color = System.Drawing.Color.FromArgb(208, 197, 227); // 底色

                RowIdx = 0;
                List<string> commentsCategorys = dt.AsEnumerable().Select(x => x.Field<string>("CommentsCategory")).Distinct().ToList();
                foreach (string commentsCategory in commentsCategorys)
                {
                    worksheet.Cells[RowIdx + 2, 1] = commentsCategory;
                    int ColumnIdx = 0;
                    foreach (string comnments in dt.AsEnumerable().Where(x => x.Field<string>("CommentsCategory") == commentsCategory).OrderBy(x => x.Field<string>("SampleStage")).Select(x => x.Field<string>("Comnments")))
                    {
                        worksheet.Cells[RowIdx + 2, ColumnIdx + 2] = comnments;
                        ColumnIdx++;
                    }
                    RowIdx++;
                }

                worksheet.Range["A2", $"A{commentsCategorys.Count + 1}"].Interior.Color = System.Drawing.Color.FromArgb(253, 233, 217); // 底色

                worksheet.Cells.EntireColumn.AutoFit();
                worksheet.Cells.EntireRow.AutoFit();

                string fileName = $"CFT Comments{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);

                result.TempFileName = fileName;
                result.Result = true;
            }
            catch (Exception ex)
            {
                throw ex;
                //result.Result = false;
                //result.ErrorMessage = ex.Message;
            }

            return result;
        }
    }
}
