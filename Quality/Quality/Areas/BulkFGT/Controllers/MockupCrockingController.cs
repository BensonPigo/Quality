

using ADOHelper.Utility.Interface;
using BusinessLogicLayer;
using BusinessLogicLayer.Service.FinalInspection;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using DatabaseObject.ViewModel.FinalInspection;
using FactoryDashBoardWeb.Helper;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Quality.Controllers;
using Sci;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace Quality.Areas.BulkFGT.Controllers
{
    public class MockupCrockingController : Controller
    {
        // GET: BulkFGT/MockupCrocking
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult DownLoad()
        {
            // 來自 Production QA P11_Detail Crocking 報表
            bool test = false;
            if (!test)
            {
                if (TempData["Model"] == null)
                {
                    return RedirectToAction("Index");
                }
            }

            MockupCrocking_ViewModel model = (MockupCrocking_ViewModel)TempData["Model"];
            TempData["Model"] = model;

            // 假設 畫面上所選傳入的 model 是第1筆
            MockupCrocking mockupCrocking = model.MockupCrocking[0];
            List<MockupCrocking_Detail> mockupCrocking_Detail = model.MockupCrocking_Detail.Where(w => w.ReportNo == mockupCrocking.ReportNo).ToList();

            #region Test data
            if (test)
            {

            }
            #endregion

            string basefileName = "MockupCrocking";
            Excel.Application excelApp = MyUtility.Excel.ConnectExcel(AppDomain.CurrentDomain.BaseDirectory + $"XLT\\{basefileName}.xltx");
            excelApp.DisplayAlerts = false;
            excelApp.Visible = test;
            Excel.Worksheet worksheet = excelApp.Sheets[1];

            // 設定表頭資料
            worksheet.Cells[4, 2] = mockupCrocking.ReportNo;
            worksheet.Cells[5, 2] = mockupCrocking.T1Subcon + "-" + mockupCrocking.Abb;

            worksheet.Cells[6, 2] = mockupCrocking.BrandID;
            worksheet.Cells[4, 6] = mockupCrocking.ReleasedDate;
            worksheet.Cells[5, 6] = mockupCrocking.TestDate;
            worksheet.Cells[6, 6] = mockupCrocking.SeasonID;

            worksheet.Cells[13, 2] = mockupCrocking.TechnicianName;
            Excel.Range cell = worksheet.Cells[12, 2];

            string picSource = mockupCrocking.SignaturePic;
            Image img = null;
            if (!MyUtility.Check.Empty(picSource))
            {
                if (System.IO.File.Exists(picSource))
                {
                    img = Image.FromFile(picSource);
                    worksheet.Shapes.AddPicture(picSource, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);
                }
            }

            #region 表身資料
            // 插入多的row
            if (mockupCrocking_Detail.Count > 0)
            {
                Excel.Range rngToInsert = worksheet.get_Range("A10:G10", Type.Missing).EntireRow;
                for (int i = 1; i < mockupCrocking_Detail.Count; i++)
                {
                    rngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                }

                Marshal.ReleaseComObject(rngToInsert);
            }

            // 塞進資料
            int start_row = 10;
            foreach (var item in mockupCrocking_Detail)
            {
                string remark = item.Remark;
                worksheet.Cells[start_row, 1] = mockupCrocking.StyleID;
                worksheet.Cells[start_row, 2] = string.IsNullOrEmpty(item.FabricColorName) ? item.FabricRefNo : item.FabricRefNo + " - " + item.FabricColorName;
                worksheet.Cells[start_row, 3] = mockupCrocking.ArtworkTypeID + "/" + item.Design + " - " + item.ArtworkColor;
                worksheet.Cells[start_row, 4] = string.IsNullOrEmpty(item.DryScale) ? string.Empty : "GRADE" + item.DryScale;
                worksheet.Cells[start_row, 5] = string.IsNullOrEmpty(item.WetScale) ? string.Empty : "GRADE" + item.WetScale;
                worksheet.Cells[start_row, 6] = item.Result;
                worksheet.Cells[start_row, 7] = item.Remark;
                worksheet.Rows[start_row].Font.Bold = false;
                worksheet.Rows[start_row].WrapText = true;
                worksheet.Rows[start_row].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // 合併儲存格無法AutoFit()因此要自己算高度
                if ((remark.Length / 20) > 1)
                {
                    worksheet.Range[$"E{start_row}", $"E{start_row}"].RowHeight = remark.Length / 20 * 16.5;
                }
                else
                {
                    worksheet.Rows[start_row].AutoFit();
                }

                start_row++;
            }
            #endregion

            #region 存檔 > 讀取MemoryStream > 下載 > 刪除
            string fileName = $"{basefileName}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
            string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileName);
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            workbook.SaveAs(filepath);
            workbook.Close();
            excelApp.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            string strPDFFileName = $"{basefileName}_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.PDF";
            if (ConvertToPDF.ExcelToPDF(filepath, strPDFFileName))
            {
                var tempFile = System.IO.Path.Combine(strPDFFileName);
                MemoryStream obj_stream = new MemoryStream(System.IO.File.ReadAllBytes(strPDFFileName));
                Response.AddHeader("Content-Disposition", $"attachment; filename={strPDFFileName}");
                Response.BinaryWrite(obj_stream.ToArray());
                obj_stream.Close();
                obj_stream.Dispose();
                Response.Flush();
                Response.End();
                System.IO.File.Delete(filepath);
                System.IO.File.Delete(strPDFFileName);
            }
            #endregion

            return null;
        }
    }
}