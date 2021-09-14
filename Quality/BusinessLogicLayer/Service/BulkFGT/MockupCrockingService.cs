using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
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
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service
{
    public class MockupCrockingService : IMockupCrockingService
    {
        private IMockupCrockingProvider _MockupCrockingProvider;
        private IMockupCrockingDetailProvider _MockupCrockingDetailProvider;
        private IStyleArtworkProvider _IStyleArtworkProvider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;

        public MockupCrocking_ViewModel GetMockupCrocking(MockupCrocking_Request MockupCrocking)
        {
            MockupCrocking.Type = "B";
            MockupCrocking_ViewModel mockupCrocking_model = new MockupCrocking_ViewModel();
            mockupCrocking_model.Request = MockupCrocking;
            try
            {
                _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
                _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(Common.ProductionDataAccessLayer);
                mockupCrocking_model = _MockupCrockingProvider.GetMockupCrocking(MockupCrocking, istop1: true).ToList().FirstOrDefault();
                if (mockupCrocking_model != null)
                {
                    mockupCrocking_model.ReportNo_Source = _MockupCrockingProvider.GetMockupCrockingReportNoList(MockupCrocking).Select(s => s.ReportNo).ToList();
                    MockupCrocking_Detail mockupCrocking_Detail = new MockupCrocking_Detail() { ReportNo = mockupCrocking_model.ReportNo };
                    mockupCrocking_model.MockupCrocking_Detail = _MockupCrockingDetailProvider.GetMockupCrocking_Detail(mockupCrocking_Detail).ToList();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mockupCrocking_model;
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

        public List<Orders> GetOrders(Orders orders)
        {
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            try
            {
                orders.Category = "B";
                return _OrdersProvider.Get(orders).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Order_Qty> GetDistinctArticle(Order_Qty order_Qty)
        {
            _OrderQtyProvider = new OrderQtyProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _OrderQtyProvider.GetDistinctArticle(order_Qty).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public Report_Result GetPDF(MockupCrocking_ViewModel mockupCrocking, bool test = false)
        {
            Report_Result result = new Report_Result();
            if (mockupCrocking == null)
            {
                result.Result = false;
                result.ErrorMessage = "Get Data Fail!";
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
                worksheet.Cells[5, 2] = mockupCrocking.T1Subcon + "-" + mockupCrocking.T1SubconAbb;

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
                    result.TempFileName = fileNamePDF;
                    result.Result = true;
                }
                else
                {
                    result.ErrorMessage = "Convert To PDF Fail";
                    result.Result = false;
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.ToString();
                result.Result = false;
            }

            return result;
        }

        public BaseResult Create(MockupCrocking_ViewModel MockupCrocking, string Mdivision, out string NewReportNo)
        {
            NewReportNo = string.Empty;
            if (MockupCrocking.MockupCrocking_Detail.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
            {
                MockupCrocking.Result = "Fail";
            }
            else
            {
                MockupCrocking.Result = "Pass";
            }

            MockupCrocking.Type = "B";
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupCrockingProvider = new MockupCrockingProvider(_ISQLDataTransaction);
            _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(_ISQLDataTransaction);
            int count;
            try
            {
                if (MockupCrocking.MockupCrocking_Detail != null || MockupCrocking.MockupCrocking_Detail.Count > 0)
                {
                    if (MockupCrocking.MockupCrocking_Detail.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                    {
                        MockupCrocking.Result = "Fail";
                    }
                    else
                    {
                        MockupCrocking.Result = "Pass";
                    }
                }
                else
                {
                    MockupCrocking.Result = string.Empty;
                }

                count = _MockupCrockingProvider.Create(MockupCrocking, Mdivision, out NewReportNo);
                if (count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Create MockupCrocking Fail. 0 Count";
                    return result;
                }

                MockupCrocking.ReportNo = NewReportNo;
                if (MockupCrocking.MockupCrocking_Detail != null)
                {
                    foreach (var MockupCrocking_Detail in MockupCrocking.MockupCrocking_Detail)
                    {
                        MockupCrocking_Detail.ReportNo = MockupCrocking.ReportNo;
                        count = _MockupCrockingDetailProvider.Create(MockupCrocking_Detail);
                        if (count == 0)
                        {
                            result.Result = false;
                            result.ErrorMessage = "Create MockupCrocking_Detail Fail. 0 Count";
                            return result;
                        }
                    }
                }

                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Create MockupCrocking Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult Update(MockupCrocking_ViewModel MockupCrocking)
        {
            if (MockupCrocking.MockupCrocking_Detail.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
            {
                MockupCrocking.Result = "Fail";
            }
            else
            {
                MockupCrocking.Result = "Pass";
            }

            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupCrockingProvider = new MockupCrockingProvider(_ISQLDataTransaction);
            _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(_ISQLDataTransaction);
            try
            {
                _MockupCrockingProvider.Update(MockupCrocking);
                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Update MockupCrocking Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult Delete(MockupCrocking_ViewModel MockupCrocking)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupCrockingProvider = new MockupCrockingProvider(_ISQLDataTransaction);
            _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(_ISQLDataTransaction);
            try
            {
                _MockupCrockingProvider.Delete(MockupCrocking);
                _MockupCrockingDetailProvider.Delete(new MockupCrocking_Detail_ViewModel() { ReportNo = MockupCrocking.ReportNo });
                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Delete MockupCrocking Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult DeleteDetail(List<MockupCrocking_Detail_ViewModel> MockupCrockingDetail)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(_ISQLDataTransaction);
            try
            {
                foreach (var MockupCrocking_Detail in MockupCrockingDetail)
                {
                    MockupCrocking_Detail.ReportNo = null;
                    _MockupCrockingDetailProvider.Delete(MockupCrocking_Detail);
                }

                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Delete MockupCrocking Detail Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public SendMail_Result FailSendMail(MockupFailMail_Request mail_Request)
        {
            _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
            string mailBody = MailTools.DataTableChangeHtml(_MockupCrockingProvider.GetMockupCrockingFailMailContentData(mail_Request.ReportNo));
            SendMail_Request sendMail_Request = new SendMail_Request();
            sendMail_Request.Subject = "Mockup Crocking – Test Fail";
            sendMail_Request.To = mail_Request.To;
            sendMail_Request.CC = mail_Request.CC;
            sendMail_Request.Body = mailBody;
            return MailTools.SendMail(sendMail_Request);
        }
    }
}
