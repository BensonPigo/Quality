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
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Web.Mvc;
using System.Windows;
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
                mockupCrocking_model.ErrorMessage = ex.Message.Replace("'", string.Empty);
                mockupCrocking_model.ReturnResult = false;
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
            catch (Exception)
            {
                return null;
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
            catch (Exception)
            {
                return null;
            }
        }

        public List<Order_Qty> GetDistinctArticle(Order_Qty order_Qty)
        {
            _OrderQtyProvider = new OrderQtyProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _OrderQtyProvider.GetDistinctArticle(order_Qty).ToList();
            }
            catch (Exception)
            {
                return null;
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
                worksheet.Cells[4, 7] = mockupCrocking.ReleasedDate;
                worksheet.Cells[5, 7] = mockupCrocking.TestDate;
                worksheet.Cells[6, 7] = mockupCrocking.SeasonID;

                worksheet.Cells[13, 2] = mockupCrocking.TechnicianName;
                Excel.Range cell = worksheet.Cells[12, 2];

                if (mockupCrocking.Signature != null)
                {
                    MemoryStream ms = new MemoryStream(mockupCrocking.Signature);                    
                    Image img = Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath;
                    if (test)
                    {
                        imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", imageName);
                    }
                    else
                    {
                        imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);
                    }

                    img.Save(imgPath);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);
                }

                Excel.Range cellBefore = worksheet.Cells[16, 1];
                if (mockupCrocking.TestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(mockupCrocking.TestBeforePicture, mockupCrocking.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: test);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBefore.Left + 5, cellBefore.Top + 5, 288, 272);
                }

                Excel.Range cellAfter = worksheet.Cells[16, 5];
                if (mockupCrocking.TestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(mockupCrocking.TestAfterPicture, mockupCrocking.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: test);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfter.Left + 5, cellAfter.Top + 5, 265, 272);              
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
                    string fabric = string.IsNullOrEmpty(item.FabricColorName) ? item.FabricRefNo : item.FabricRefNo + " - " + item.FabricColorName;
                    string artwork = string.IsNullOrEmpty(mockupCrocking.ArtworkTypeID) ? item.Design + " - " + item.ArtworkColorName : mockupCrocking.ArtworkTypeID + "/" + item.Design + " - " + item.ArtworkColorName;
                    string remark = item.Remark;
                    worksheet.Cells[start_row, 1] = mockupCrocking.StyleID;
                    worksheet.Cells[start_row, 2] = fabric;
                    worksheet.Cells[start_row, 3] = artwork;
                    worksheet.Cells[start_row, 5] = string.IsNullOrEmpty(item.DryScale) ? string.Empty : "GRADE" + item.DryScale;
                    worksheet.Cells[start_row, 6] = string.IsNullOrEmpty(item.WetScale) ? string.Empty : "GRADE" + item.WetScale;
                    worksheet.Cells[start_row, 7] = item.Result;
                    worksheet.Cells[start_row, 8] = item.Remark;
                    worksheet.Rows[start_row].Font.Bold = false;
                    worksheet.Rows[start_row].WrapText = true;
                    worksheet.Rows[start_row].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    int maxLength = fabric.Length > remark.Length ? fabric.Length : remark.Length;
                    maxLength = maxLength > artwork.Length ? maxLength : artwork.Length;
                    worksheet.Range[$"A{start_row}", $"H{start_row}"].RowHeight = ((maxLength / 20) + 1) * 16.5;

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
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
                result.Result = false;
            }

            return result;
        }

        public BaseResult Create(MockupCrocking_ViewModel MockupCrocking, string Mdivision, string userid, out string NewReportNo)
        {
            NewReportNo = string.Empty;
            MockupCrocking.Type = "B";
            MockupCrocking.AddName = userid;
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupCrockingProvider = new MockupCrockingProvider(_ISQLDataTransaction);
            _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(_ISQLDataTransaction);
            int count;
            try
            {
                if (MockupCrocking.MockupCrocking_Detail != null && MockupCrocking.MockupCrocking_Detail.Count > 0)
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
                        MockupCrocking_Detail.EditName = userid;
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

        public BaseResult Update(MockupCrocking_ViewModel MockupCrocking, string userid)
        {
            if (MockupCrocking.MockupCrocking_Detail != null && MockupCrocking.MockupCrocking_Detail.Count > 0)
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

            MockupCrocking.EditName = userid;
            if (MockupCrocking.MockupCrocking_Detail != null)
            {
                foreach (var item in MockupCrocking.MockupCrocking_Detail)
                {
                    item.EditName = userid;
                }
            }
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupCrockingProvider = new MockupCrockingProvider(_ISQLDataTransaction);

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

        public SendMail_Result FailSendMail(MockupFailMail_Request mail_Request )
        {
            _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
            string mailBody = MailTools.DataTableChangeHtml(_MockupCrockingProvider.GetMockupCrockingFailMailContentData(mail_Request.ReportNo), out AlternateView plainView);
            MockupCrocking_ViewModel model = GetMockupCrocking(new MockupCrocking_Request { ReportNo = mail_Request.ReportNo });
            Report_Result baseResult = GetPDF(model);
            string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", baseResult.TempFileName) : string.Empty;
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = "Mockup Crocking – Test Fail",
                To = mail_Request.To,
                CC = mail_Request.CC,
                Body = mailBody,
                alternateView = plainView,
                FileonServer = new List<string> { FileName },
            };
            return MailTools.SendMail(sendMail_Request);
        }
    }
}
