using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using ClosedXML.Excel;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Web.Mvc;

namespace BusinessLogicLayer.Service
{
    public class MockupWashService : IMockupWashService
    {
        private IMockupWashProvider _MockupWashProvider;
        private IStyleBOAProvider _IStyleBOAProvider;
        private IStyleArtworkProvider _IStyleArtworkProvider;
        private IDropDownListProvider _DropDownListProvider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;
        private IInspectionTypeProvider _InspectionTypeProvider;
        private QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;
        private MailToolsService _MailService;

        public MockupWash_ViewModel GetMockupWash(MockupWash_Request MockupWash)
        {
            MockupWash.Type = "B";
            MockupWash_ViewModel mockupWash_model = new MockupWash_ViewModel();
            mockupWash_model.Request = MockupWash;
            try
            {
                _MockupWashProvider = new MockupWashProvider(Common.ProductionDataAccessLayer);
                mockupWash_model = _MockupWashProvider.GetMockupWash(MockupWash, istop1: true).ToList().FirstOrDefault();
                if (mockupWash_model != null)
                {
                    mockupWash_model.ReportNo_Source = _MockupWashProvider.GetMockupWashReportNoList(MockupWash).Select(s => s.ReportNo).ToList();
                    MockupWash_Detail mockupWash_Detail = new MockupWash_Detail() { ReportNo = mockupWash_model.ReportNo };
                    mockupWash_model.MockupWash_Detail = _MockupWashProvider.GetMockupWash_Detail(mockupWash_Detail).ToList();

                    mockupWash_model.TestingMethod_Source = new List<SelectListItem>();
                    _DropDownListProvider = new DropDownListProvider(Common.ProductionDataAccessLayer);
                    DropDownList downList = new DropDownList() { Type = "PMS_MockupWashMethod" };
                    List<DropDownList> dropDowns = _DropDownListProvider.Get(downList).ToList();
                    foreach (var item in dropDowns)
                    {
                        mockupWash_model.TestingMethod_Source.Add(new SelectListItem() { Value = item.ID, Text = item.Description });
                    }


                    mockupWash_model.MailSubject = $"Mockup Wash /{mockupWash_model.POID}/" +
                        $"{mockupWash_model.StyleID}/" +
                        $"{mockupWash_model.Article}/" +
                        $"{mockupWash_model.Result}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                }
            }
            catch (Exception ex)
            {
                mockupWash_model.ErrorMessage = ex.Message.Replace("'", string.Empty);
                mockupWash_model.ReturnResult = false;
            }

            return mockupWash_model;
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

        public List<SelectListItem> GetTestingMethod()
        {
            List<SelectListItem> selectListItems = new List<SelectListItem>();
            _DropDownListProvider = new DropDownListProvider(Common.ProductionDataAccessLayer);
            DropDownList downList = new DropDownList() { Type = "PMS_MockupWashMethod" };
            List<DropDownList> dropDowns = _DropDownListProvider.Get(downList).ToList();
            foreach (var item in dropDowns)
            {
                selectListItems.Add(new SelectListItem() { Value = item.ID, Text = item.Description });
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
        public Report_Result GetPDF(MockupWash_ViewModel mockupWash, bool test = false, string AssignedFineName = "")
        {
            Report_Result result = new Report_Result();
            string tmpName = string.Empty;
            if (mockupWash == null)
            {
                result.Result = false;
                result.ErrorMessage = "Get Data Fail!";
                return result;
            }

            try
            {
                string basePath = test ? AppDomain.CurrentDomain.BaseDirectory : System.Web.HttpContext.Current.Server.MapPath("~/");

                string xltPath = Path.Combine(basePath, "XLT", mockupWash.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER") ? "MockupWash2.xltx" : "MockupWash.xltx");
                string tmpPath = Path.Combine(basePath, "TMP");

                if (!Directory.Exists(tmpPath)) Directory.CreateDirectory(tmpPath);

                tmpName = $"Mockup Wash _{mockupWash.POID}_{mockupWash.StyleID}_{mockupWash.Article}_{mockupWash.Result}_{DateTime.Now:yyyyMMddHHmmss}";

                if (!File.Exists(xltPath)) throw new FileNotFoundException("Template not found", xltPath);

                using (var workbook = new XLWorkbook(xltPath))
                {
                    bool isHT = mockupWash.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER");
                    int haveHTrow = isHT ? 6 : 0;
                    var worksheet = workbook.Worksheet(1);

                    worksheet.Cell(2, 1).Value = mockupWash.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER")
                        ? "Wash TEST (HEAT TRANSFER)"
                        : "Wash TEST";

                    worksheet.Cell(4, 2).Value = mockupWash.ReportNo;
                    worksheet.Cell(5, 2).Value = $"{mockupWash.T1Subcon} - {mockupWash.T1SubconAbb}";
                    worksheet.Cell(6, 2).Value = $"{mockupWash.T2Supplier} - {mockupWash.T2SupplierAbb}";
                    worksheet.Cell(7, 2).Value = mockupWash.MethodDescription;

                    worksheet.Cell(4, 6).Value = mockupWash.ReleasedDate;
                    worksheet.Cell(5, 6).Value = mockupWash.TestDate;
                    worksheet.Cell(6, 6).Value = mockupWash.ReceivedDate;
                    worksheet.Cell(7, 6).Value = mockupWash.SeasonID;

                    if (isHT)
                    {
                        worksheet.Cell(10, 2).Value = mockupWash.HTPlate;
                        worksheet.Cell(11, 2).Value = mockupWash.HTFlim;
                        worksheet.Cell(12, 2).Value = mockupWash.HTTime;
                        worksheet.Cell(13, 2).Value = mockupWash.HTPressure;

                        worksheet.Cell(10, 6).Value = mockupWash.HTPellOff;
                        worksheet.Cell(11, 6).Value = mockupWash.HT2ndPressnoreverse;
                        worksheet.Cell(12, 6).Value = mockupWash.HT2ndPressreversed;
                        worksheet.Cell(13, 6).Value = mockupWash.HTCoolingTime;
                    }

                    worksheet.Cell(13 + haveHTrow, 2).Value = mockupWash.TechnicianName;
                    AddImageToWorksheet(worksheet, mockupWash.Signature, 12 + haveHTrow, 2, 100, 24);

                    AddImageToWorksheet(worksheet, mockupWash.TestBeforePicture, 18 + haveHTrow, 1, 288, 272);
                    AddImageToWorksheet(worksheet, mockupWash.TestAfterPicture, 18 + haveHTrow, 4, 265, 272);

                    int startRow = 10 + haveHTrow;
                    if (mockupWash.MockupWash_Detail.Count > 0)
                    {
                        for (int i = 1; i < mockupWash.MockupWash_Detail.Count; i++)
                        {
                            // 1. 複製第 10 列
                            var rowToCopy = worksheet.Row(10);

                            // 2. 插入一列，將第 10 和第 11 列之間騰出空間
                            worksheet.Row(11).InsertRowsAbove(1);

                            // 3. 複製內容與格式到新插入的第 11 列
                            var newRow = worksheet.Row(11);
                            rowToCopy.CopyTo(newRow);
                        }

                        foreach (var item in mockupWash.MockupWash_Detail)
                        {
                            string fabric = string.IsNullOrEmpty(item.FabricColorName) ? item.FabricRefNo : item.FabricRefNo + " - " + item.FabricColorName;
                            string artwork = string.IsNullOrEmpty(mockupWash.ArtworkTypeID) ? item.Design + " - " + item.ArtworkColorName : mockupWash.ArtworkTypeID + "/" + item.Design + " - " + item.ArtworkColorName;
                            worksheet.Cell(startRow, 1).Value = mockupWash.StyleID;
                            worksheet.Cell(startRow, 2).Value = fabric;
                            worksheet.Cell(startRow, 3).Value = artwork;
                            worksheet.Cell(startRow, 4).Value = item.Result;
                            worksheet.Cell(startRow, 5).Value = item.Remark;

                            int maxLength = fabric.Length > item.Remark.Length ? fabric.Length : item.Remark.Length;
                            maxLength = maxLength > artwork.Length ? maxLength : artwork.Length;
                            worksheet.Range(worksheet.Cell(startRow, 1), worksheet.Cell(startRow, 5)).Row(startRow).Style.Font.Bold = false;
                            worksheet.Range(worksheet.Cell(startRow, 1), worksheet.Cell(startRow, 5)).Row(startRow).Style.Alignment.WrapText = true;                            
                            worksheet.Row(startRow).Height = ((maxLength / 20) + 1) * 16.5;
                            startRow++;
                        }
                    }

                    // ISP20230792
                    if ((mockupWash.HTPlate > 0)
                        || (mockupWash.HTFlim > 0)
                        || (mockupWash.HTTime > 0)
                        || (mockupWash.HTPressure > 0)
                        || !string.IsNullOrEmpty(mockupWash.HTPellOff)
                        || (mockupWash.HT2ndPressnoreverse > 0)
                        || (mockupWash.HT2ndPressreversed > 0)
                        || !string.IsNullOrEmpty(mockupWash.HTCoolingTime))
                    {
                        worksheet.Row(startRow).InsertRowsAbove(5);
                        int aRow = startRow;
                        workbook.Worksheet(2).Range("A1:J5").CopyTo(worksheet.Range($"A{aRow}:J{aRow + 4}"));
                        worksheet.Cell(aRow + 1, 3).Value = mockupWash.HTPlate;
                        worksheet.Cell(aRow + 2, 3).Value = mockupWash.HTFlim;
                        worksheet.Cell(aRow + 3, 3).Value = mockupWash.HTTime;
                        worksheet.Cell(aRow + 4, 3).Value = mockupWash.HTPressure;
                        worksheet.Cell(aRow + 1, 8).Value = mockupWash.HTPellOff;
                        worksheet.Cell(aRow + 2, 8).Value = mockupWash.HT2ndPressnoreverse;
                        worksheet.Cell(aRow + 3, 8).Value = mockupWash.HT2ndPressreversed;
                        worksheet.Cell(aRow + 4, 8).Value = mockupWash.HTCoolingTime;
                    }

                    // Excel 合併 + 塞資料
                    #region Title
                    string FactoryNameEN = _MockupWashProvider.GetFactoryNameEN(mockupWash.ReportNo, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
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

                    workbook.Worksheet(2).Delete();
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

        public BaseResult Create(MockupWash_ViewModel MockupWash, string Mdivision, string userid, out string NewReportNo)
        {
            NewReportNo = string.Empty;
            MockupWash.Type = "B";
            MockupWash.AddName = userid;
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupWashProvider = new MockupWashProvider(_ISQLDataTransaction);
            int count;
            try
            {
                if (MockupWash.MockupWash_Detail != null && MockupWash.MockupWash_Detail.Count > 0)
                {
                    if (MockupWash.MockupWash_Detail.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                    {
                        MockupWash.Result = "Fail";
                    }
                    else
                    {
                        MockupWash.Result = "Pass";
                    }
                }
                else
                {
                    MockupWash.Result = string.Empty;
                }

                count = _MockupWashProvider.CreateMockupWash(MockupWash, Mdivision, out NewReportNo);
                if (count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Create MockupWash Fail. 0 Count";
                    return result;
                }

                MockupWash.ReportNo = NewReportNo;
                foreach (var MockupWash_Detail in MockupWash.MockupWash_Detail)
                {
                    MockupWash_Detail.EditName = userid;
                    MockupWash_Detail.ReportNo = MockupWash.ReportNo;
                    count = _MockupWashProvider.CreateDetail(MockupWash_Detail);
                    if (count == 0)
                    {
                        result.Result = false;
                        result.ErrorMessage = "Create MockupWash_Detail Fail. 0 Count";
                        return result;
                    }
                }

                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Create MockupWash Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult Update(MockupWash_ViewModel MockupWash, string userid)
        {
            if (MockupWash.MockupWash_Detail != null && MockupWash.MockupWash_Detail.Count > 0)
            {
                if (MockupWash.MockupWash_Detail.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                {
                    MockupWash.Result = "Fail";
                }
                else
                {
                    MockupWash.Result = "Pass";
                }
            }
            else
            {
                MockupWash.Result = string.Empty;
            }

            MockupWash.EditName = userid;
            if (MockupWash.MockupWash_Detail != null)
            {
                foreach (var item in MockupWash.MockupWash_Detail)
                {
                    item.EditName = userid;
                }
            }
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupWashProvider = new MockupWashProvider(_ISQLDataTransaction);
            try
            {
                _MockupWashProvider.UpdateMockupWash(MockupWash);
                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Update MockupWash Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult Delete(MockupWash_ViewModel MockupWash)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupWashProvider = new MockupWashProvider(_ISQLDataTransaction);
            try
            {
                _MockupWashProvider.DeleteMockupWash(MockupWash);
                _MockupWashProvider.DeleteDetail(new MockupWash_Detail_ViewModel() { ReportNo = MockupWash.ReportNo });
                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Delete MockupWash Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult DeleteDetail(List<MockupWash_Detail_ViewModel> MockupWashDetail)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupWashProvider = new MockupWashProvider(_ISQLDataTransaction);
            try
            {
                foreach (var MockupWash_Detail in MockupWashDetail)
                {
                    MockupWash_Detail.ReportNo = null;
                    _MockupWashProvider.DeleteDetail(MockupWash_Detail);
                }

                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Delete MockupWash Detail Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public SendMail_Result SendMail(MockupFailMail_Request mail_Request)
        {
            _MockupWashProvider = new MockupWashProvider(Common.ProductionDataAccessLayer);
            System.Data.DataTable dt = _MockupWashProvider.GetMockupWashFailMailContentData(mail_Request.ReportNo);

            MockupWash_ViewModel model = GetMockupWash(new MockupWash_Request { ReportNo = mail_Request.ReportNo });

            string name = $"Mockup Wash _{model.POID}_" +
                    $"{model.StyleID}_" +
                    $"{model.Article}_" +
                    $"{model.Result}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
            Report_Result baseResult = GetPDF(model, AssignedFineName: name);
            string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", baseResult.TempFileName) : string.Empty;
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Mockup Wash /{model.POID}/" +
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
                AICommentType = "Mockup Wash Test",
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
