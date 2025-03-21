using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using ClosedXML.Excel;
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
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service
{
    public class MockupCrockingService : IMockupCrockingService
    {
        private IMockupCrockingProvider _MockupCrockingProvider;
        private IStyleArtworkProvider _IStyleArtworkProvider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;
        private MailToolsService _MailService;

        public MockupCrocking_ViewModel GetMockupCrocking(MockupCrocking_Request MockupCrocking)
        {
            MockupCrocking.Type = "B";
            MockupCrocking_ViewModel model = new MockupCrocking_ViewModel();
            model.Request = MockupCrocking;
            try
            {
                _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);

                model = _MockupCrockingProvider.GetMockupCrocking(MockupCrocking, istop1: true).ToList().FirstOrDefault();
                if (model != null)
                {
                    model.ReportNo_Source = _MockupCrockingProvider.GetMockupCrockingReportNoList(MockupCrocking).Select(s => s.ReportNo).ToList();
                    MockupCrocking_Detail mockupCrocking_Detail = new MockupCrocking_Detail() { ReportNo = model.ReportNo };
                    model.MockupCrocking_Detail = _MockupCrockingProvider.GetMockupCrocking_Detail(mockupCrocking_Detail).ToList();

                    model.MailSubject = $"Mockup Crocking /{model.POID}/" +
                    $"{model.StyleID}/" +
                    $"{model.Article}/" +
                    $"{model.Result}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                }

            }
            catch (Exception ex)
            {
                model.ErrorMessage = ex.Message.Replace("'", string.Empty);
                model.ReturnResult = false;
            }

            return model;
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

        public Report_Result GetPDF(MockupCrocking_ViewModel mockupCrocking, string AssignedFineName = "")
        {
            Report_Result result = new Report_Result();
            if (mockupCrocking == null)
            {
                result.Result = false;
                result.ErrorMessage = "Get Data Fail!";
                return result;
            }

            string tmpName = string.Empty;

            try
            {
                string basePath = System.Web.HttpContext.Current.Server.MapPath("~/");

                string templatePath = Path.Combine(basePath, "XLT", "MockupCrocking.xltx");
                string tmpPath = Path.Combine(basePath, "TMP");


                tmpName = $"Mockup Crocking _{mockupCrocking.POID}_{mockupCrocking.StyleID}_{mockupCrocking.Article}_{mockupCrocking.Result}_{DateTime.Now:yyyyMMddHHmmss}";
                if (!string.IsNullOrEmpty(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }
                // 去除非法字元
                tmpName = FileNameHelper.SanitizeFileName(tmpName);

                if (!File.Exists(templatePath)) throw new FileNotFoundException("Template not found", templatePath);

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    worksheet.Cell(4, 2).Value = mockupCrocking.ReportNo;
                    worksheet.Cell(5, 2).Value = $"{mockupCrocking.T1Subcon}-{mockupCrocking.T1SubconAbb}";
                    worksheet.Cell(6, 2).Value = mockupCrocking.BrandID;
                    worksheet.Cell(4, 7).Value = mockupCrocking.ReleasedDate;
                    worksheet.Cell(5, 7).Value = mockupCrocking.TestDate;
                    worksheet.Cell(6, 7).Value = mockupCrocking.SeasonID;
                    worksheet.Cell(13, 2).Value = mockupCrocking.TechnicianName;

                    AddImageToWorksheet(worksheet, mockupCrocking.Signature, 12, 2, 100, 24);
                    AddImageToWorksheet(worksheet, mockupCrocking.TestBeforePicture, 16, 1, 288, 272);
                    AddImageToWorksheet(worksheet, mockupCrocking.TestAfterPicture, 16, 5, 265, 272);

                    // 表身筆數處理，複製儲存格
                    if (mockupCrocking.MockupCrocking_Detail.Count > 1)
                    {
                        for (int i = 1; i < mockupCrocking.MockupCrocking_Detail.Count; i++)
                        {
                            // 1. 複製第 10 列
                            var rowToCopy = worksheet.Row(10);

                            // 2. 插入一列，將第 10 和第 11 列之間騰出空間
                            worksheet.Row(11).InsertRowsAbove(1);

                            // 3. 複製內容與格式到新插入的第 11 列
                            var newRow = worksheet.Row(11);
                            rowToCopy.CopyTo(newRow);
                        }
                    }

                    int startRow = 10;
                    foreach (var item in mockupCrocking.MockupCrocking_Detail)
                    {
                        string fabric = string.IsNullOrEmpty(item.FabricColorName) ? item.FabricRefNo : $"{item.FabricRefNo} - {item.FabricColorName}";
                        string artwork = string.IsNullOrEmpty(mockupCrocking.ArtworkTypeID) ? $"{item.Design} - {item.ArtworkColorName}" : $"{mockupCrocking.ArtworkTypeID}/{item.Design} - {item.ArtworkColorName}";

                        worksheet.Cell(startRow, 1).Value = mockupCrocking.StyleID;
                        worksheet.Cell(startRow, 2).Value = fabric;
                        worksheet.Cell(startRow, 3).Value = artwork;
                        worksheet.Cell(startRow, 5).Value = string.IsNullOrEmpty(item.DryScale) ? string.Empty : $"GRADE {item.DryScale}";
                        worksheet.Cell(startRow, 6).Value = string.IsNullOrEmpty(item.WetScale) ? string.Empty : $"GRADE {item.WetScale}";
                        worksheet.Cell(startRow, 7).Value = item.Result;
                        worksheet.Cell(startRow, 8).Value = item.Remark;

                        worksheet.Row(startRow).AdjustToContents();
                        startRow++;
                    }

                    tmpName = RemoveInvalidFileNameChars(tmpName);

                    string xlsxPath = Path.Combine(tmpPath, tmpName + ".xlsx");
                    string pdfPath = Path.Combine(tmpPath, tmpName + ".pdf");


                    // Excel 合併 + 塞資料
                    #region Title
                    string FactoryNameEN = _MockupCrockingProvider.GetFactoryNameEN(mockupCrocking.ReportNo, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                    // 1. 插入一列
                    worksheet.Row(2).InsertRowsAbove(1);

                    // 2. 合併欄位
                    worksheet.Range("A1:H1").Merge();
                    worksheet.Range("A2:H2").Merge();
                    // 設置字體樣式
                    var mergedCell = worksheet.Cell("A2");
                    mergedCell.Value = FactoryNameEN;
                    mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                    mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                    mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    mergedCell.Style.Font.Bold = true;
                    mergedCell.Style.Font.Italic = false;
                    #endregion

                    workbook.SaveAs(xlsxPath);


                    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                    //officeService.ConvertExcelToPdf(xlsxPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                    //ConvertToPDF.ExcelToPDF(xlsxPath, pdfPath);
                    //result.TempFileName = tmpName + ".pdf";
                    result.TempFileName = tmpName + ".xlsx";
                    result.Result = true;
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.Result = false;
            }

            return result;
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

        private string RemoveInvalidFileNameChars(string input)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                input = input.Replace(c.ToString(), "");
            }
            return input;
        }
        public BaseResult Create(MockupCrocking_ViewModel model, string Mdivision, string userid, out string NewReportNo)
        {
            NewReportNo = string.Empty;
            model.Type = "B";
            model.AddName = userid;
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupCrockingProvider = new MockupCrockingProvider(_ISQLDataTransaction);
            int count;
            try
            {
                if (model.MockupCrocking_Detail != null && model.MockupCrocking_Detail.Count > 0)
                {
                    if (model.MockupCrocking_Detail.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                    {
                        model.Result = "Fail";
                    }
                    else
                    {
                        model.Result = "Pass";
                    }
                }
                else
                {
                    model.Result = string.Empty;
                }

                count = _MockupCrockingProvider.CreateMockupCrocking(model, Mdivision, out NewReportNo);
                if (count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Create MockupCrocking Fail. 0 Count";
                    return result;
                }

                model.ReportNo = NewReportNo;
                if (model.MockupCrocking_Detail != null)
                {
                    foreach (var MockupCrocking_Detail in model.MockupCrocking_Detail)
                    {
                        MockupCrocking_Detail.EditName = userid;
                        MockupCrocking_Detail.ReportNo = model.ReportNo;
                        count = _MockupCrockingProvider.CreateDetail(MockupCrocking_Detail);
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
                _MockupCrockingProvider.UpdateMockupCrocking(MockupCrocking);
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
            try
            {
                _MockupCrockingProvider.DeleteMockupCrocking(MockupCrocking);
                _MockupCrockingProvider.DeleteDetail(new MockupCrocking_Detail_ViewModel() { ReportNo = MockupCrocking.ReportNo });
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
            _MockupCrockingProvider = new MockupCrockingProvider(_ISQLDataTransaction);
            try
            {
                foreach (var MockupCrocking_Detail in MockupCrockingDetail)
                {
                    MockupCrocking_Detail.ReportNo = null;
                    _MockupCrockingProvider.DeleteDetail(MockupCrocking_Detail);
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

        public SendMail_Result SendMail(MockupFailMail_Request mail_Request)
        {
            _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
            System.Data.DataTable dt = _MockupCrockingProvider.GetMockupCrockingFailMailContentData(mail_Request.ReportNo);
            
            MockupCrocking_ViewModel model = GetMockupCrocking(new MockupCrocking_Request { ReportNo = mail_Request.ReportNo });

            string name = $"Mockup Crocking _{model.POID}_" +
                $"{model.StyleID}_" +
                $"{model.Article}_" +
                $"{model.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";


            Report_Result baseResult = GetPDF(model,AssignedFineName: name);
            string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", baseResult.TempFileName) : string.Empty;
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Mockup Crocking /{model.POID}/" +
                $"{model.StyleID}/" +
                $"{model.Article}/" +
                $"{model.Result}/" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                To = mail_Request.To,
                CC = mail_Request.CC,
                //Body = mailBody,
                //alternateView = plainView,
                FileonServer = new List<string> { FileName },
                FileUploader = mail_Request.Files,
                IsShowAIComment = true,
                AICommentType = "Mockup Crocking Test",
                StyleID = model.StyleID,
                SeasonID = model.SeasonID,
                BrandID = model.BrandID,
            };

            if (!string.IsNullOrEmpty(mail_Request.Subject))
            {
                sendMail_Request.Subject = mail_Request.Subject;
            }

            _MailService = new MailToolsService();
            string comment = _MailService.GetAICommet(sendMail_Request);
            string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
            string mailBody = MailTools.DataTableChangeHtml(dt, comment, buyReadyDate, mail_Request.Body, out AlternateView plainView);

            sendMail_Request.Body = mailBody;
            sendMail_Request.alternateView = plainView;

            return MailTools.SendMail(sendMail_Request);
        }
    }
}
