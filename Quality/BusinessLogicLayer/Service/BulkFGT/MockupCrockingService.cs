using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;
using System.IO;
using Library;

namespace BusinessLogicLayer.Service
{
    public class MockupCrockingService : IMockupCrockingService
    {
        private IMockupCrockingProvider _MockupCrockingProvider;
        private IMockupCrockingDetailProvider _MockupCrockingDetailProvider;

        public MockupCrocking_ViewModel GetMockupCrocking(MockupCrocking MockupCrocking)
        {
            MockupCrocking_ViewModel model = new MockupCrocking_ViewModel();
            try
            {
                _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
                _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(Common.ProductionDataAccessLayer);
                model.MockupCrocking = _MockupCrockingProvider.GetMockupCrocking(MockupCrocking).ToList();
                foreach (var item in model.MockupCrocking)
                {
                    MockupCrocking_Detail mockupCrocking_Detail = new MockupCrocking_Detail() { ReportNo = item.ReportNo };
                    item.MockupCrocking_Detail = _MockupCrockingDetailProvider.GetMockupCrocking_Detail(mockupCrocking_Detail).ToList();
                }

                model.Result = true;
            }
            catch (System.Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        // 單獨更新圖檔欄位
        public MockupCrocking_ViewModel UpdatePicture(MockupCrocking MockupCrocking)
        {
            MockupCrocking_ViewModel model = new MockupCrocking_ViewModel();
            _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
            try
            {
                int updateCT = _MockupCrockingProvider.UpdatePicture(MockupCrocking);
                if (updateCT == 0)
                {
                    model.Result = false;
                    model.ErrorMessage = "Not found Crocking Data!";
                }
                else
                {
                    model.Result = true;
                }
            }
            catch (System.Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.ToString();
            }

            return model;
        }

        public MockupCrocking_ViewModel GetExcel(MockupCrocking MockupCrocking)
        {
            MockupCrocking_ViewModel result = new MockupCrocking_ViewModel();
            try
            {
                result = GetMockupCrocking(MockupCrocking);

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                }

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                }

                _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);


                MockupCrocking mockupCrocking = result.MockupCrocking[0];
                List<MockupCrocking_Detail> mockupCrocking_Detail = mockupCrocking.MockupCrocking_Detail;
                string basefileName = "MockupCrocking";
                Excel.Application excelApp = MyUtility.Excel.ConnectExcel(System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx");
                excelApp.DisplayAlerts = false;
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
                if (!MyUtility.Check.Empty(picSource))
                {
                    if (System.IO.File.Exists(picSource))
                    {
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

                string fileName = $"{basefileName}{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);
                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);

                string fileNamePDF = $"{basefileName}{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.PDF";
                if (ConvertToPDF.ExcelToPDF(filepath, fileNamePDF))
                {
                    result.TempFileName = fileNamePDF;
                    result.Result = true;
                }
                else
                {
                    result.ErrorMessage = "ConvertToPDF Fail";
                    result.Result = false;
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public MockupCrocking_ViewModel Create(MockupCrocking_ViewModel MockupCrocking)
        {
            MockupCrocking_ViewModel model = new MockupCrocking_ViewModel();
            _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
            _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(Common.ProductionDataAccessLayer);
            int insertCt;
            try
            {
                insertCt = _MockupCrockingProvider.Create(MockupCrocking.MockupCrocking[0]);
                foreach (var MockupCrocking_Detail in MockupCrocking.MockupCrocking[0].MockupCrocking_Detail)
                {
                    _MockupCrockingDetailProvider.Create(MockupCrocking_Detail);
                }

                if (insertCt == 0)
                {
                    model.Result = false;
                    model.ErrorMessage = "Insert 0 count Data!";

                }
                else
                {
                    model.Result = true;
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return model;
        }
    }
}
