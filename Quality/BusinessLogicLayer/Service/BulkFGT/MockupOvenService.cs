using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using Microsoft.Office.Interop.Excel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Transactions;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service
{
    public class MockupOvenService : IMockupOvenService
    {
        private IMockupOvenProvider _MockupOvenProvider;
        private IMockupOvenDetailProvider _MockupOvenDetailProvider;
        private IStyleBOAProvider _IStyleBOAProvider;
        private IStyleArtworkProvider _IStyleArtworkProvider;

        public MockupOven_ViewModel GetMockupOven(MockupOven_Request MockupOven)
        {
            MockupOven_ViewModel mockupOven_model = new MockupOven_ViewModel();
            try
            {
                _MockupOvenProvider = new MockupOvenProvider(Common.ProductionDataAccessLayer);
                _MockupOvenDetailProvider = new MockupOvenDetailProvider(Common.ProductionDataAccessLayer);
                mockupOven_model = _MockupOvenProvider.GetMockupOven(MockupOven, istop1: true).ToList().First();
                mockupOven_model.ReportNo_Source = _MockupOvenProvider.GetMockupOvenReportNoList(MockupOven).Select(s => s.ReportNo).ToList();
                MockupOven_Detail mockupOven_Detail = new MockupOven_Detail() { ReportNo = mockupOven_model.ReportNo };
                mockupOven_model.MockupOven_Detail = _MockupOvenDetailProvider.GetMockupOven_Detail(mockupOven_Detail).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mockupOven_model;
        }

        public List<SelectListItem> GetArtworkTypeID(StyleArtwork_Request Request)
        {
            List<SelectListItem> selectListItems = new List<SelectListItem>();
            try
            {
                _IStyleArtworkProvider = new StyleArtworkProvider(Common.ProductionDataAccessLayer);
                var ArtworkTypeID = _IStyleArtworkProvider.GetArtworkTypeID(Request).ToList();
                foreach (var item in ArtworkTypeID)
                {
                    selectListItems.Add(new SelectListItem { Value = item.ArtworkTypeID, Text = item.ArtworkTypeID });
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return selectListItems;
        }

        public List<SelectListItem> GetAccessoryRefNo(AccessoryRefNo_Request Request)
        {
            Request.MtlTypeID = "HEAT TRANS";
            List<SelectListItem> selectListItems = new List<SelectListItem>();
            try
            {
                _IStyleBOAProvider = new StyleBOAProvider(Common.ProductionDataAccessLayer);
                var AccessoryRefNos = _IStyleBOAProvider.GetAccessoryRefNo(Request).ToList();
                foreach (var item in AccessoryRefNos)
                {
                    selectListItems.Add(new SelectListItem { Value = item.Refno, Text = item.Refno });
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return selectListItems;
        }

        public MockupOven_ViewModel GetPDF(MockupOven_ViewModel mockupOven, bool test = false)
        {
            MockupOven_ViewModel result = new MockupOven_ViewModel();
            if (mockupOven == null)
            {
                result.ReportResult = false;
                result.ReportErrorMessage = "Get Data Fail!";
                return result;
            }


            try
            {
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

                _MockupOvenProvider = new MockupOvenProvider(Common.ProductionDataAccessLayer);


                var mockupOven_Detail = mockupOven.MockupOven_Detail;
                string basefileName = "MockupOven";
                string openfilepath;
                if (test)
                {
                    openfilepath = "C:\\Git\\Quality\\Quality\\Quality\\bin\\XLT\\MockupOven.xltx";
                }
                else
                {
                    openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";
                }

                Application excelApp = MyUtility.Excel.ConnectExcel(openfilepath);



                excelApp.DisplayAlerts = false;
                Worksheet worksheet = excelApp.Sheets[1];

                bool haveHT = mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER");
                int htRow = 6;
                int haveHTrow = haveHT ? htRow : 0;

                // 設定表頭資料
                worksheet.Cells[4, 2] = mockupOven.ReportNo;
                worksheet.Cells[5, 2] = mockupOven.T1Subcon + "-" + mockupOven.T1SubconAbb;
                worksheet.Cells[6, 2] = mockupOven.T2Supplier + "-" + mockupOven.T2SupplierAbb;
                worksheet.Cells[7, 2] = mockupOven.BrandID;
                worksheet.Cells[8, 2] = $"5.14 color migration test({mockupOven.TestTemperature} degree @ {mockupOven.TestTime} hours)";

                worksheet.Cells[4, 6] = mockupOven.ReleasedDate;
                worksheet.Cells[5, 6] = mockupOven.TestDate;
                worksheet.Cells[6, 6] = mockupOven.SeasonID;

                if (haveHT)
                {
                    worksheet.Cells[10, 2] = mockupOven.HTPlate;
                    worksheet.Cells[11, 2] = mockupOven.HTFlim;
                    worksheet.Cells[12, 2] = mockupOven.HTTime;
                    worksheet.Cells[13, 2] = mockupOven.HTPressure;
                    worksheet.Cells[10, 6] = mockupOven.HTPellOff;
                    worksheet.Cells[11, 6] = mockupOven.HT2ndPressnoreverse;
                    worksheet.Cells[12, 6] = mockupOven.HT2ndPressreversed;
                    worksheet.Cells[13, 6] = mockupOven.HTCoolingTime;
                }
                else
                {
                    for (int i = 0; i < htRow; i++)
                    {
                        worksheet.Rows[9].Delete(XlDeleteShiftDirection.xlShiftUp);
                    }
                }

                worksheet.Cells[13 + haveHTrow, 2] = mockupOven.TechnicianName;

                Range cell = worksheet.Cells[12 + haveHTrow, 2];

                string picSource = mockupOven.SignaturePic;
                if (!MyUtility.Check.Empty(picSource))
                {
                    if (File.Exists(picSource))
                    {
                        worksheet.Shapes.AddPicture(picSource, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);
                    }
                }

                #region 表身資料
                // 插入多的row
                if (mockupOven_Detail.Count > 0)
                {
                    Range rngToInsert = worksheet.get_Range($"A{10 + haveHTrow}:G{10 + haveHTrow}", Type.Missing).EntireRow;
                    for (int i = 1; i < mockupOven_Detail.Count; i++)
                    {
                        rngToInsert.Insert(XlInsertShiftDirection.xlShiftDown);
                        worksheet.get_Range(string.Format("E{0}:G{0}", (10 + haveHTrow + i - 1).ToString())).Merge(false);
                    }

                    Marshal.ReleaseComObject(rngToInsert);
                }

                // 塞進資料
                int start_row = 10 + haveHTrow;
                foreach (var item in mockupOven_Detail)
                {
                    string remark = item.Remark;
                    string fabric = string.IsNullOrEmpty(item.FabricColorName) ? item.FabricRefNo : item.FabricRefNo + " - " + item.FabricColorName;
                    string artwork = mockupOven.ArtworkTypeID + "/" + item.Design + " - " + item.ArtworkColor;
                    worksheet.Cells[start_row, 1] = mockupOven.StyleID;
                    worksheet.Cells[start_row, 2] = fabric;
                    worksheet.Cells[start_row, 3] = artwork;
                    worksheet.Cells[start_row, 4] = item.Result;
                    worksheet.Cells[start_row, 5] = item.Remark;
                    worksheet.Rows[start_row].Font.Bold = false;
                    worksheet.Rows[start_row].WrapText = true;
                    worksheet.Rows[start_row].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    // 合併儲存格無法AutoFit()因此要自己算高度
                    if (fabric.Length > remark.Length || artwork.Length > remark.Length)
                    {
                        worksheet.Rows[start_row].AutoFit();
                    }
                    else
                    {
                        worksheet.Range[$"E{start_row}", $"E{start_row}"].RowHeight = ((remark.Length / 20) + 1) * 16.5;
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


                Workbook workbook = excelApp.ActiveWorkbook;
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
                    result.ReportErrorMessage = "Convert To PDF Fail";
                    result.ReportResult = false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public MockupOven_ViewModel Create(MockupOven_ViewModel MockupOven)
        {
            MockupOven_ViewModel model = new MockupOven_ViewModel();
            _MockupOvenProvider = new MockupOvenProvider(Common.ProductionDataAccessLayer);
            _MockupOvenDetailProvider = new MockupOvenDetailProvider(Common.ProductionDataAccessLayer);
            int insertCt;
            try
            {
                using (TransactionScope scope = new TransactionScope())
                {
                    insertCt = _MockupOvenProvider.Create(MockupOven);
                    if (insertCt == 0)
                    {
                        return model;
                    }

                    foreach (var MockupOven_Detail in MockupOven.MockupOven_Detail)
                    {
                        insertCt = _MockupOvenDetailProvider.Create(MockupOven_Detail);
                        if (insertCt == 0)
                        {
                            return model;
                        }
                    }

                    scope.Complete();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return model;
        }

        public MockupOven_ViewModel Update(MockupOven_ViewModel MockupOven)
        {
            MockupOven_ViewModel model = new MockupOven_ViewModel();
            _MockupOvenProvider = new MockupOvenProvider(Common.ProductionDataAccessLayer);
            _MockupOvenDetailProvider = new MockupOvenDetailProvider(Common.ProductionDataAccessLayer);
            int insertCt;
            try
            {
                using (TransactionScope scope = new TransactionScope())
                {
                    insertCt = _MockupOvenProvider.Update(MockupOven);
                    if (insertCt == 0)
                    {
                        return model;
                    }

                    foreach (var MockupOven_Detail in MockupOven.MockupOven_Detail)
                    {
                        insertCt = _MockupOvenDetailProvider.Update(MockupOven_Detail);
                        if (insertCt == 0)
                        {
                            return model;
                        }
                    }

                    scope.Complete();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return model;
        }

        public MockupOven_ViewModel Delete(MockupOven_ViewModel MockupOven)
        {
            MockupOven_ViewModel model = new MockupOven_ViewModel();
            _MockupOvenProvider = new MockupOvenProvider(Common.ProductionDataAccessLayer);
            _MockupOvenDetailProvider = new MockupOvenDetailProvider(Common.ProductionDataAccessLayer);
            int insertCt;
            try
            {
                using (TransactionScope scope = new TransactionScope())
                {
                    insertCt = _MockupOvenProvider.Delete(MockupOven);
                    if (insertCt == 0)
                    {
                        return model;
                    }

                    foreach (var MockupOven_Detail in MockupOven.MockupOven_Detail)
                    {
                        insertCt = _MockupOvenDetailProvider.Delete(MockupOven_Detail);
                        if (insertCt == 0)
                        {
                            return model;
                        }
                    }

                    scope.Complete();
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
