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
                        model.SampleStageList = res.Select(o => o.SampleStage).Distinct().ToList();
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
            Excel.Application excelApp = null;
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
                excelApp = MyUtility.Excel.ConnectExcel(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\CFT Comments.xltx");

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

                result.TempFileName = fileName;
                result.Result = true;
            }
            catch (Exception ex)
            {
                throw ex;
                //result.Result = false;
                //result.ErrorMessage = ex.Message;
            }
            finally
            {
                MyUtility.Excel.KillExcelProcess(excelApp);
            }

            return result;
        }

        /// <summary>
        /// 透過  Microsoft.Office.Interop.Excel 生成Excel，並下載至暫存路徑，新版本格式
        /// </summary>
        /// <param name="Req"></param>
        /// <returns>暫存檔路徑</returns>
        public CFTComments_ViewModel GetExcel2(CFTComments_ViewModel Req)
        {
            string TempTilePath = string.Empty;
            CFTComments_ViewModel result = new CFTComments_ViewModel();
            DataTable dt = new DataTable();
            Excel.Application excelApp = null;

            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);
 
                // 取得Datatable
                List<CFTComments_Result> datas =_CFTCommentsProvider.Get_CFT_OrderComments2(Req).ToList();
                List<string> SampleStageList = Req.SampleStageList;

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                }

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                }

                //if (datas.Count == 0)
                //{
                //    result.TempFileName = string.Empty;
                //    result.Result = false;
                //    return result;
                //}

                // 開啟excel app
                excelApp = MyUtility.Excel.ConnectExcel(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\CFT Comment Report.xltx");
                excelApp.Visible = false;
                Excel.Worksheet worksheet;

                List<string> sampleStages = datas.Select(x => x.SampleStage).Distinct().ToList();

                // 複製分頁：表身幾筆，就幾個sheet
                if (sampleStages.Count > 1)
                {
                    for (int j = 1; j < sampleStages.Count; j++)
                    {
                        worksheet = (Excel.Worksheet)excelApp.ActiveWorkbook.Worksheets[j];

                        worksheet.Copy(worksheet);
                    }
                }

                for (int i = 1; i <= sampleStages.Count; i++)
                {
                    worksheet = (Excel.Worksheet)excelApp.ActiveWorkbook.Worksheets[i];
                    worksheet.Name = sampleStages[i-1];
                }

                //開始填資料
                for (int i = 1; i <= sampleStages.Count; i++)
                {
                    string sampleStage = sampleStages[i-1];
                    var sameData = datas.Where(o => o.SampleStage == sampleStage).ToList();

                    Excel.Worksheet currenSheet = excelApp.ActiveWorkbook.Worksheets[i];

                    currenSheet.Cells[3, 1] = sameData.FirstOrDefault().StyleID;
                    currenSheet.Cells[4, 2] = sameData.FirstOrDefault().SampleStage;
                    currenSheet.Cells[7, 2] = sameData.FirstOrDefault().Article;
                    currenSheet.Cells[8, 2] = sameData.FirstOrDefault().SizeCode;

                    currenSheet.Cells[1, 16] = "Season : " + sameData.FirstOrDefault().SeasonID;
                    currenSheet.Cells[2, 16] = "Date : " + DateTime.Now.ToString("yyyy/MM/dd");
                    currenSheet.Cells[3, 14] = "Released by: " + sameData.FirstOrDefault().Name;

                    currenSheet.Cells[10, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "SAMPLE MEASUREMENT").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "SAMPLE MEASUREMENT").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[12, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "SAMPLE FITTING").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "SAMPLE FITTING").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[13, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "MATERIAL").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "MATERIAL").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[16, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "ACCESSORY").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "ACCESSORY").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[18, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "WORKMANSHIP").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "WORKMANSHIP").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[24, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "ARTWORK").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "ARTWORK").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[27, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "HANGTAG").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "HANGTAG").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[28, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "LABELLING AND PACKAGING").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "LABELLING AND PACKAGING").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[31, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "SAMPLE WEIGHT(SIZE SET ONLY)").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "SAMPLE WEIGHT(SIZE SET ONLY)").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[32, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "RESULT(SEALING AND SIZE SET ONLY)").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "RESULT(SEALING AND SIZE SET ONLY)").FirstOrDefault().Comnments : string.Empty;
                    currenSheet.Cells[33, 2] = sameData.Where(o => o.CommentsCategory.ToUpper() == "REMARK").Any() ? sameData.Where(o => o.CommentsCategory.ToUpper() == "REMARK").FirstOrDefault().Comnments : string.Empty;

                    var sizeList = (sameData.FirstOrDefault().SizeCode != null ? sameData.FirstOrDefault().SizeCode.Split(',').OrderBy(o => o).ToList() : new List<string>());

                    int ctn = 0;
                    foreach (var size in sizeList)
                    {
                        currenSheet.Cells[42, 3 + ctn] = size;
                        ctn++;
                    }
                }

                string fileName = $"CFT Comments{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();

                result.TempFileName = fileName;
                result.Result = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                MyUtility.Excel.KillExcelProcess(excelApp);
            }

            return result;
        }
    }
}
