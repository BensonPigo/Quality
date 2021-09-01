using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Transactions;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service
{
    public class MockupCrockingService : IMockupCrockingService
    {
        private IMockupCrockingProvider _MockupCrockingProvider;
        private IMockupCrockingDetailProvider _MockupCrockingDetailProvider;

        public MockupCrockings_ViewModel GetMockupCrocking(MockupCrocking_Request MockupCrocking)
        {
            MockupCrockings_ViewModel model = new MockupCrockings_ViewModel();
            try
            {
                _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
                _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(Common.ProductionDataAccessLayer);
                model.MockupCrocking = _MockupCrockingProvider.GetMockupCrocking(MockupCrocking).ToList();
                model.ReportNos = new List<string>();
                foreach (var item in model.MockupCrocking)
                {
                    model.ReportNos.Add(item.ReportNo);
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

        public MockupCrocking_ViewModel GetExcel(string ReportNo)
        {
            bool test = false;
            MockupCrocking_ViewModel result = new MockupCrocking_ViewModel();

            var oneReportNo = GetMockupCrocking(new MockupCrocking_Request() { ReportNo = ReportNo });
            if (oneReportNo == null)
            {
                result.ReportResult = false;
                result.ErrorMessage = "Get Data Fail!";
                return result;
            }

            if (oneReportNo.MockupCrocking.Count == 0)
            {
                result.ReportResult = false;
                result.ErrorMessage = "Data Not found!";
                return result;
            }

            try
            {
                var mockupCrocking = oneReportNo.MockupCrocking[0];
                if (!test)
                {
                    if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                    {
                        System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                    }

                    if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                    {
                        System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                    }
                }

                _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);


                var mockupCrocking_Detail = mockupCrocking.MockupCrocking_Detail;
                string basefileName = "MockupCrocking";
                string openfilepath;
                if (test)
                {
                    openfilepath = "C:\\Git\\Quality\\Quality\\Quality\\bin\\XLT\\MockupCrocking.xltx";
                }
                else
                {
                    openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";
                }

                Excel.Application excelApp = MyUtility.Excel.ConnectExcel(openfilepath);



                excelApp.DisplayAlerts = false;
                Excel.Worksheet worksheet = excelApp.Sheets[1];

                // 設定表頭資料
                worksheet.Cells[4, 2] = mockupCrocking.ReportNo;
                worksheet.Cells[5, 2] = mockupCrocking.T1SubconName;

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

                string fileName = $"{basefileName}{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";
                string filexlsx = fileName + ".xlsx";
                string fileNamePDF = fileName + ".pdf";

                string filepath;
                string filepathpdf;
                if (test)
                {
                    filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", filexlsx);
                    filepathpdf = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileNamePDF);
                }
                else
                {
                    filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);
                    filepathpdf = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF);
                }


                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);
                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);


                if (ConvertToPDF.ExcelToPDF(filepath, filepathpdf))
                {
                    result.TempFileName = filepathpdf;
                    result.ReportResult = true;
                }
                else
                {
                    result.ErrorMessage = "Convert To PDF Fail";
                    result.ReportResult = false;
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public MockupCrockings_ViewModel Create(MockupCrockings_ViewModel MockupCrocking)
        {
            MockupCrockings_ViewModel model = new MockupCrockings_ViewModel();
            _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
            _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(Common.ProductionDataAccessLayer);
            int insertCt;
            try
            {
                using (TransactionScope scope = new TransactionScope())
                {
                    insertCt = _MockupCrockingProvider.Create(MockupCrocking.MockupCrocking[0]);
                    if (insertCt == 0)
                    {
                        model.Result = false;
                        model.ErrorMessage = "Insert MockupCrocking Fail!";
                        return model;
                    }

                    foreach (var MockupCrocking_Detail in MockupCrocking.MockupCrocking[0].MockupCrocking_Detail)
                    {
                        insertCt = _MockupCrockingDetailProvider.Create(MockupCrocking_Detail);
                        if (insertCt == 0)
                        {
                            model.Result = false;
                            model.ErrorMessage = "Insert MockupCrocking_Detail Fail!";
                            return model;
                        }
                    }

                    scope.Complete();
                }

                model.Result = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return model;
        }

        public MockupCrockingScale GetScale()
        {
            MockupCrockingScale mockupCrockingScale = new MockupCrockingScale
            {
                DryScale = new List<string>() { "1", "1-2", "2", "2-3", "3", "3-4", "4", "4-5", "5" },
                WetScale = new List<string>() { "1", "1-2", "2", "2-3", "3", "3-4", "4", "4-5", "5" },
            };

            return mockupCrockingScale;
        }
    }
}
