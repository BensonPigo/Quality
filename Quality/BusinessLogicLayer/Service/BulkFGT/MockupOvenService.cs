using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using ClosedXML.Excel;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service
{
    public class MockupOvenService : IMockupOvenService
    {
        private IMockupOvenProvider _MockupOvenProvider;
        private IStyleBOAProvider _IStyleBOAProvider;
        private IStyleArtworkProvider _IStyleArtworkProvider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;
        private IScaleProvider _ScaleProvider;
        private IInspectionTypeProvider _InspectionTypeProvider;
        private QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;
        private MailToolsService _MailService;

        private string IsTest = ConfigurationManager.AppSettings["IsTest"];

        public MockupOven_ViewModel GetMockupOven(MockupOven_Request MockupOven)
        {
            MockupOven.Type = "B";
            MockupOven_ViewModel mockupOven_model = new MockupOven_ViewModel();
            mockupOven_model.Request = MockupOven;
            try
            {
                _MockupOvenProvider = new MockupOvenProvider(Common.ProductionDataAccessLayer);
                mockupOven_model = _MockupOvenProvider.GetMockupOven(MockupOven, istop1: true).ToList().FirstOrDefault();
                if (mockupOven_model != null)
                {
                    mockupOven_model.ScaleID_Source = GetScale();
                    mockupOven_model.ReportNo_Source = _MockupOvenProvider.GetMockupOvenReportNoList(MockupOven).Select(s => s.ReportNo).ToList();
                    MockupOven_Detail mockupOven_Detail = new MockupOven_Detail() { ReportNo = mockupOven_model.ReportNo };
                    mockupOven_model.MockupOven_Detail = _MockupOvenProvider.GetMockupOven_Detail(mockupOven_Detail).ToList();

                    mockupOven_model.MailSubject = $"Mockup Oven /{mockupOven_model.POID}/" +
                    $"{mockupOven_model.StyleID}/" +
                    $"{mockupOven_model.Article}/" +
                    $"{mockupOven_model.Result}/" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                }

            }
            catch (Exception ex)
            {
                mockupOven_model.ErrorMessage = ex.Message.Replace("'", string.Empty);
                mockupOven_model.ReturnResult = false;
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
            catch (Exception)
            {
                return null;
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
            catch (Exception)
            {
                return null;
            }

            return selectListItems;
        }

        public List<SelectListItem> GetScale()
        {
            _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);
            List<SelectListItem> selectListItems = new List<SelectListItem>();
            try
            {
                foreach (string item in _ScaleProvider.Get().Select(s => s.ID))
                {
                    selectListItems.Add(new SelectListItem { Value = item, Text = item });
                }
            }
            catch (Exception)
            {

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

        public Report_Result GetPDF(MockupOven_ViewModel mockupOven, string AssignedFineName = "")
        {
            Report_Result result = new Report_Result();
            if (mockupOven == null)
            {
                result.Result = false;
                result.ErrorMessage = "Get Data Fail!";
                return result;
            }

            string tmpName = string.Empty;

            try
            {
                string basePath = IsTest.ToLower() == "true" ? AppDomain.CurrentDomain.BaseDirectory : System.Web.HttpContext.Current.Server.MapPath("~/");

                string xltPath = Path.Combine(basePath, "XLT", mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER") ? "MockupOven2.xltx" : "MockupOven.xltx");
                string tmpPath = Path.Combine(basePath, "TMP");

                if (!Directory.Exists(tmpPath)) Directory.CreateDirectory(tmpPath);

                tmpName = $"Mockup Oven _{mockupOven.POID}_{mockupOven.StyleID}_{mockupOven.Article}_{mockupOven.Result}_{DateTime.Now:yyyyMMddHHmmss}";

                if (!File.Exists(xltPath)) throw new FileNotFoundException("Template not found", xltPath);

                using (var workbook = new XLWorkbook(xltPath))
                {
                    var worksheet = workbook.Worksheet(1);

                    worksheet.Cell(2, 1).Value = mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER")
                        ? "COLOR MIGRATION TEST (Oven) (HEAT TRANSFER)"
                        : "COLOR MIGRATION TEST (Oven)";

                    worksheet.Cell(4, 2).Value = mockupOven.ReportNo;
                    worksheet.Cell(5, 2).Value = $"{mockupOven.T1Subcon} - {mockupOven.T1SubconAbb}";
                    worksheet.Cell(6, 2).Value = $"{mockupOven.T2Supplier} - {mockupOven.T2SupplierAbb}";
                    worksheet.Cell(7, 2).Value = mockupOven.BrandID;
                    worksheet.Cell(8, 2).Value = $"5.14 color migration test ({mockupOven.TestTemperature} degree @ {mockupOven.TestTime} hours)";

                    worksheet.Cell(4, 8).Value = mockupOven.ReleasedDate;
                    worksheet.Cell(5, 8).Value = mockupOven.TestDate;
                    worksheet.Cell(6, 8).Value = mockupOven.SeasonID;

                    if (mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER"))
                    {
                        worksheet.Cell(10, 2).Value = mockupOven.HTPlate;
                        worksheet.Cell(11, 2).Value = mockupOven.HTFlim;
                        worksheet.Cell(12, 2).Value = mockupOven.HTTime;
                        worksheet.Cell(13, 2).Value = mockupOven.HTPressure;

                        worksheet.Cell(10, 8).Value = mockupOven.HTPellOff;
                        worksheet.Cell(11, 8).Value = mockupOven.HT2ndPressnoreverse;
                        worksheet.Cell(12, 8).Value = mockupOven.HT2ndPressreversed;
                        worksheet.Cell(13, 8).Value = mockupOven.HTCoolingTime;
                    }

                    worksheet.Cell(13 + (mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER") ? 6 : 0), 2).Value = mockupOven.TechnicianName;
                    AddImageToWorksheet(worksheet, mockupOven.Signature, 12 + (mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER") ? 6 : 0), 2, 100, 24);

                    AddImageToWorksheet(worksheet, mockupOven.TestBeforePicture, 16 + (mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER") ? 6 : 0), 2, 288, 272);
                    AddImageToWorksheet(worksheet, mockupOven.TestAfterPicture, 16 + (mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER") ? 6 : 0), 8, 265, 272);

                    if (mockupOven.MockupOven_Detail.Count > 0)
                    {
                        for (int i = 1; i < mockupOven.MockupOven_Detail.Count; i++)
                        {
                            // 1. 複製第 10 列
                            var rowToCopy = worksheet.Row(10);

                            // 2. 插入一列，將第 10 和第 11 列之間騰出空間
                            worksheet.Row(11).InsertRowsAbove(1);

                            // 3. 複製內容與格式到新插入的第 11 列
                            var newRow = worksheet.Row(11);
                            rowToCopy.CopyTo(newRow);
                        }

                        int startRow = 10 + (mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER") ? 6 : 0);

                        foreach (var item in mockupOven.MockupOven_Detail)
                        {
                            worksheet.Cell(startRow, 1).Value = mockupOven.StyleID;
                            worksheet.Cell(startRow, 2).Value = string.IsNullOrEmpty(item.FabricColorName) ? item.FabricRefNo : $"{item.FabricRefNo} - {item.FabricColorName}";
                            worksheet.Cell(startRow, 3).Value = string.IsNullOrEmpty(mockupOven.ArtworkTypeID) ? $"{item.Design} - {item.ArtworkColorName}" : $"{mockupOven.ArtworkTypeID}/{item.Design} - {item.ArtworkColorName}";
                            worksheet.Cell(startRow, 5).Value = item.ChangeScale;
                            worksheet.Cell(startRow, 6).Value = item.ResultChange;
                            worksheet.Cell(startRow, 7).Value = item.StainingScale;
                            worksheet.Cell(startRow, 8).Value = item.ResultStain;
                            worksheet.Cell(startRow, 9).Value = item.Remark;

                            worksheet.Row(startRow).AdjustToContents();
                            startRow++;
                        }
                    }

                    tmpName = RemoveInvalidFileNameChars(tmpName);

                    string filePath = Path.Combine(tmpPath, $"{tmpName}.xlsx");
                    string pdfPath = Path.Combine(tmpPath, $"{tmpName}.pdf");

                    workbook.SaveAs(filePath);

                    LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                    officeService.ConvertExcelToPdf(filePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                    result.TempFileName = $"{tmpName}.pdf";
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

        public BaseResult Create(MockupOven_ViewModel MockupOven, string Mdivision, string userid, out string NewReportNo)
        {
            NewReportNo = string.Empty;
            MockupOven.Type = "B";
            MockupOven.AddName = userid;
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupOvenProvider = new MockupOvenProvider(_ISQLDataTransaction);
            int count;
            try
            {
                if (MockupOven.MockupOven_Detail != null && MockupOven.MockupOven_Detail.Count > 0)
                {
                    if (MockupOven.MockupOven_Detail.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                    {
                        MockupOven.Result = "Fail";
                    }
                    else
                    {
                        MockupOven.Result = "Pass";
                    }
                }
                else
                {
                    MockupOven.Result = string.Empty;
                }

                count = _MockupOvenProvider.CreateMockupOven(MockupOven, Mdivision, out NewReportNo);
                if (count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Create MockupOven Fail. 0 Count";
                    return result;
                }

                MockupOven.ReportNo = NewReportNo;
                foreach (var MockupOven_Detail in MockupOven.MockupOven_Detail)
                {
                    MockupOven_Detail.EditName = userid;
                    MockupOven_Detail.ReportNo = MockupOven.ReportNo;
                    count = _MockupOvenProvider.CreateDetail(MockupOven_Detail);
                    if (count == 0)
                    {
                        result.Result = false;
                        result.ErrorMessage = "Create MockupOven_Detail Fail. 0 Count";
                        return result;
                    }
                }

                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Create MockupOven Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult Update(MockupOven_ViewModel MockupOven, string userid)
        {
            if (MockupOven.MockupOven_Detail != null && MockupOven.MockupOven_Detail.Count > 0)
            {
                if (MockupOven.MockupOven_Detail.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                {
                    MockupOven.Result = "Fail";
                }
                else
                {
                    MockupOven.Result = "Pass";
                }
            }
            else
            {
                MockupOven.Result = string.Empty;
            }

            MockupOven.EditName = userid;
            if (MockupOven.MockupOven_Detail != null)
            {
                foreach (var item in MockupOven.MockupOven_Detail)
                {
                    item.EditName = userid;
                }
            }
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupOvenProvider = new MockupOvenProvider(_ISQLDataTransaction);
            try
            {
                _MockupOvenProvider.UpdateMockupOven(MockupOven);
                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Update MockupOven Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult Delete(MockupOven_ViewModel MockupOven)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupOvenProvider = new MockupOvenProvider(_ISQLDataTransaction);
            try
            {
                _MockupOvenProvider.DeleteMockupOven(MockupOven);
                _MockupOvenProvider.DeleteDetail(new MockupOven_Detail_ViewModel() { ReportNo = MockupOven.ReportNo });
                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Delete MockupOven Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult DeleteDetail(List<MockupOven_Detail_ViewModel> MockupOvenDetail)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupOvenProvider = new MockupOvenProvider(_ISQLDataTransaction);
            try
            {
                foreach (var MockupOven_Detail in MockupOvenDetail)
                {
                    MockupOven_Detail.ReportNo = null;
                    _MockupOvenProvider.DeleteDetail(MockupOven_Detail);
                }

                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Delete MockupOven Detail Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public SendMail_Result SendMail(MockupFailMail_Request mail_Request)
        {
            _MockupOvenProvider = new MockupOvenProvider(Common.ProductionDataAccessLayer);
            System.Data.DataTable dt = _MockupOvenProvider.GetMockupOvenFailMailContentData(mail_Request.ReportNo);
            MockupOven_ViewModel model = GetMockupOven(new MockupOven_Request { ReportNo = mail_Request.ReportNo });

            string name = $"Mockup Oven _{model.POID}_" +
                $"{model.StyleID}_" +
                $"{model.Article}_" +
                $"{model.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
            
            Report_Result baseResult = GetPDF(model, name);
            string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", baseResult.TempFileName) : string.Empty;
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Mockup Oven /{model.POID}/" +
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
                AICommentType = "Mockup Oven Test",
                StyleID = model.StyleID,
                SeasonID = model.SeasonID,
                BrandID = model.BrandID,
            };

            if (!string.IsNullOrEmpty(mail_Request.Subject))
            {
                sendMail_Request.Subject= mail_Request.Subject;
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
