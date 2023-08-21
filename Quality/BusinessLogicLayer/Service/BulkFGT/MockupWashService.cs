using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using Microsoft.Office.Interop.Excel;
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
using System.Windows;

namespace BusinessLogicLayer.Service
{
    public class MockupWashService : IMockupWashService
    {
        private IMockupWashProvider _MockupWashProvider;
        private IMockupWashDetailProvider _MockupWashDetailProvider;
        private IStyleBOAProvider _IStyleBOAProvider;
        private IStyleArtworkProvider _IStyleArtworkProvider;
        private IDropDownListProvider _DropDownListProvider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;
        private IInspectionTypeProvider _InspectionTypeProvider;
        private MailToolsService _MailService;

        public MockupWash_ViewModel GetMockupWash(MockupWash_Request MockupWash)
        {
            MockupWash.Type = "B";
            MockupWash_ViewModel mockupWash_model = new MockupWash_ViewModel();
            mockupWash_model.Request = MockupWash;
            try
            {
                _MockupWashProvider = new MockupWashProvider(Common.ProductionDataAccessLayer);
                _MockupWashDetailProvider = new MockupWashDetailProvider(Common.ProductionDataAccessLayer);
                mockupWash_model = _MockupWashProvider.GetMockupWash(MockupWash, istop1: true).ToList().FirstOrDefault();
                if (mockupWash_model != null)
                {
                    mockupWash_model.ReportNo_Source = _MockupWashProvider.GetMockupWashReportNoList(MockupWash).Select(s => s.ReportNo).ToList();
                    MockupWash_Detail mockupWash_Detail = new MockupWash_Detail() { ReportNo = mockupWash_model.ReportNo };
                    mockupWash_model.MockupWash_Detail = _MockupWashDetailProvider.GetMockupWash_Detail(mockupWash_Detail).ToList();

                    mockupWash_model.TestingMethod_Source = new List<SelectListItem>();
                    _DropDownListProvider = new DropDownListProvider(Common.ProductionDataAccessLayer);
                    DropDownList downList = new DropDownList() { Type = "PMS_MockupWashMethod" };
                    List<DropDownList> dropDowns = _DropDownListProvider.Get(downList).ToList();
                    foreach (var item in dropDowns)
                    {
                        mockupWash_model.TestingMethod_Source.Add(new SelectListItem() { Value = item.ID, Text = item.Description });
                    }
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

        public Report_Result GetPDF(MockupWash_ViewModel mockupWash, bool test = false)
        {
            Report_Result result = new Report_Result();
            if (mockupWash == null)
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

                _MockupWashProvider = new MockupWashProvider(Common.ProductionDataAccessLayer);
                _InspectionTypeProvider = new InspectionTypeProvider(Common.ProductionDataAccessLayer);

                List<InspectionType> InspectionTypes = _InspectionTypeProvider.Get_InspectionType("MockupWash", "Bulk", mockupWash.BrandID).ToList();
                mockupWash.Requirements = InspectionTypes.Select(x => x.Comment).ToList();
                var mockupWash_Detail = mockupWash.MockupWash_Detail;
                bool haveHT = mockupWash.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER");
                string basefileName = haveHT ? "MockupWash2" : "MockupWash";
                string openfilepath;
                if (test)
                {
                    openfilepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "XLT", $"{basefileName}.xltx");
                }
                else
                {
                    openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";
                }

                Application excelApp = MyUtility.Excel.ConnectExcel(openfilepath);
                excelApp.DisplayAlerts = false;
                Worksheet worksheet = excelApp.Sheets[1];
                Worksheet worksheet2 = excelApp.Sheets[2];
                int htRow = 6;
                int haveHTrow = haveHT ? htRow : 0;

                // 設定表頭資料
                worksheet.Cells[4, 2] = mockupWash.ReportNo;
                worksheet.Cells[5, 2] = mockupWash.T1Subcon + "-" + mockupWash.T1SubconAbb; ;
                worksheet.Cells[6, 2] = mockupWash.T2Supplier + "-" + mockupWash.T2SupplierAbb; ;

                worksheet.Cells[7, 2] = mockupWash.MethodDescription;
                worksheet.Cells[4, 6] = mockupWash.ReleasedDate;
                worksheet.Cells[5, 6] = mockupWash.TestDate;
                worksheet.Cells[6, 6] = mockupWash.ReceivedDate;
                worksheet.Cells[7, 6] = mockupWash.SeasonID;

                if (haveHT)
                {
                    worksheet.Cells[10, 2] = mockupWash.HTPlate;
                    worksheet.Cells[11, 2] = mockupWash.HTFlim;
                    worksheet.Cells[12, 2] = mockupWash.HTTime;
                    worksheet.Cells[13, 2] = mockupWash.HTPressure;
                    worksheet.Cells[10, 6] = mockupWash.HTPellOff;
                    worksheet.Cells[11, 6] = mockupWash.HT2ndPressnoreverse;
                    worksheet.Cells[12, 6] = mockupWash.HT2ndPressreversed;
                    worksheet.Cells[13, 6] = mockupWash.HTCoolingTime;
                }

                worksheet.Cells[13 + haveHTrow, 2] = mockupWash.TechnicianName;

                Range cell = worksheet.Cells[12 + haveHTrow, 2];

                if (mockupWash.Signature != null)
                {
                    MemoryStream ms = new MemoryStream(mockupWash.Signature);
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

                Range cellBefore = worksheet.Cells[16 + haveHTrow, 1];
                if (mockupWash.TestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(mockupWash.TestBeforePicture, mockupWash.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: test);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBefore.Left, cellBefore.Top, cell.Width, cell.Height);
                }

                Range cellAfter = worksheet.Cells[16 + haveHTrow, 4];
                if (mockupWash.TestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(mockupWash.TestAfterPicture, mockupWash.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: test);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfter.Left, cellAfter.Top, cellAfter.Width, cellAfter.Height);
                }

                #region 表身資料
                // 插入多的row
                if (mockupWash_Detail.Count > 0)
                {
                    Range rngToInsert = worksheet.get_Range($"A{10 + haveHTrow}:J{10 + haveHTrow}", Type.Missing).EntireRow;
                    for (int i = 1; i < mockupWash_Detail.Count; i++)
                    {
                        rngToInsert.Insert(XlInsertShiftDirection.xlShiftDown);
                        worksheet.get_Range(string.Format("E{0}:G{0}", (10 + haveHTrow + i - 1).ToString())).Merge(false);
                    }

                    worksheet.get_Range(string.Format("H{0}:J{1}", (10 + haveHTrow).ToString(), (10 + haveHTrow + mockupWash_Detail.Count - 1).ToString())).Merge(false);
                    Marshal.ReleaseComObject(rngToInsert);
                }

                // ISP20230792
                if ((mockupWash.HTPlate.HasValue && mockupWash.HTPlate.Value > 0)
                    || (mockupWash.HTFlim.HasValue && mockupWash.HTFlim.Value > 0)
                    || (mockupWash.HTTime.HasValue && mockupWash.HTTime.Value > 0)
                    || (mockupWash.HTPressure.HasValue && mockupWash.HTPressure.Value > 0)
                    || !string.IsNullOrEmpty(mockupWash.HTPellOff) 
                    || (mockupWash.HT2ndPressnoreverse.HasValue && mockupWash.HT2ndPressnoreverse.Value > 0)
                    || (mockupWash.HT2ndPressreversed.HasValue && mockupWash.HT2ndPressreversed.Value > 0)
                    || !string.IsNullOrEmpty(mockupWash.HTCoolingTime))
                {
                    int aRow = 11 + haveHTrow + mockupWash_Detail.Count - 1;
                    Range rngToCopy = worksheet2.get_Range($"A1:J5", Type.Missing).EntireRow;
                    Range rngToInsert = worksheet.get_Range($"A{aRow}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上

                    worksheet.Cells[aRow + 1, 3] = mockupWash.HTPlate.HasValue ? mockupWash.HTPlate.Value : 0;
                    worksheet.Cells[aRow + 2, 3] = mockupWash.HTFlim.HasValue ? mockupWash.HTFlim.Value : 0;
                    worksheet.Cells[aRow + 3, 3] = mockupWash.HTTime.HasValue ? mockupWash.HTTime.Value : 0;
                    worksheet.Cells[aRow + 4, 3] = mockupWash.HTPressure.HasValue ? mockupWash.HTPressure.Value : 0;
                    worksheet.Cells[aRow + 1, 8] = mockupWash.HTPellOff;
                    worksheet.Cells[aRow + 2, 8] = mockupWash.HT2ndPressnoreverse.HasValue ? mockupWash.HT2ndPressnoreverse.Value : 0;
                    worksheet.Cells[aRow + 3, 8] = mockupWash.HT2ndPressreversed.HasValue ? mockupWash.HT2ndPressreversed.Value : 0;
                    worksheet.Cells[aRow + 4, 8] = mockupWash.HTCoolingTime;
                }

                // 塞進資料
                int start_row = 10 + haveHTrow;
                worksheet.Cells[start_row, 8] = mockupWash.Requirements.JoinToString(Environment.NewLine);
                foreach (var item in mockupWash_Detail)
                {
                    string remark = item.Remark;
                    string fabric = string.IsNullOrEmpty(item.FabricColorName) ? item.FabricRefNo : item.FabricRefNo + " - " + item.FabricColorName;
                    string artwork = string.IsNullOrEmpty(mockupWash.ArtworkTypeID) ? item.Design + " - " + item.ArtworkColorName : mockupWash.ArtworkTypeID + "/" + item.Design + " - " + item.ArtworkColorName;
                    worksheet.Cells[start_row, 1] = mockupWash.StyleID;
                    worksheet.Cells[start_row, 2] = fabric;
                    worksheet.Cells[start_row, 3] = artwork;
                    worksheet.Cells[start_row, 4] = item.Result;
                    worksheet.Cells[start_row, 5] = item.Remark;
                    worksheet.Rows[start_row].Font.Bold = false;
                    worksheet.Rows[start_row].WrapText = true;
                    worksheet.Rows[start_row].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    int maxLength = fabric.Length > remark.Length ? fabric.Length : remark.Length;
                    maxLength = maxLength > artwork.Length ? maxLength : artwork.Length;
                    worksheet.Range[$"A{start_row}", $"J{start_row}"].RowHeight = ((maxLength / 20) + 1) * 16.5;

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
                worksheet2.Visible = XlSheetVisibility.xlSheetHidden;
                workbook.SaveAs(filepath);
                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(worksheet2);
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

        public BaseResult Create(MockupWash_ViewModel MockupWash, string Mdivision, string userid, out string NewReportNo)
        {
            NewReportNo = string.Empty;
            MockupWash.Type = "B";
            MockupWash.AddName = userid;
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupWashProvider = new MockupWashProvider(_ISQLDataTransaction);
            _MockupWashDetailProvider = new MockupWashDetailProvider(_ISQLDataTransaction);
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

                count = _MockupWashProvider.Create(MockupWash, Mdivision, out NewReportNo);
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
                    count = _MockupWashDetailProvider.Create(MockupWash_Detail);
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
            _MockupWashDetailProvider = new MockupWashDetailProvider(_ISQLDataTransaction);
            try
            {
                _MockupWashProvider.Update(MockupWash);
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
            _MockupWashDetailProvider = new MockupWashDetailProvider(_ISQLDataTransaction);
            try
            {
                _MockupWashProvider.Delete(MockupWash);
                _MockupWashDetailProvider.Delete(new MockupWash_Detail_ViewModel() { ReportNo = MockupWash.ReportNo });
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
            _MockupWashDetailProvider = new MockupWashDetailProvider(_ISQLDataTransaction);
            try
            {
                foreach (var MockupWash_Detail in MockupWashDetail)
                {
                    MockupWash_Detail.ReportNo = null;
                    _MockupWashDetailProvider.Delete(MockupWash_Detail);
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

        public SendMail_Result FailSendMail(MockupFailMail_Request mail_Request)
        {
            _MockupWashProvider = new MockupWashProvider(Common.ProductionDataAccessLayer);
            System.Data.DataTable dt = _MockupWashProvider.GetMockupWashFailMailContentData(mail_Request.ReportNo);
            MockupWash_ViewModel model = GetMockupWash(new MockupWash_Request { ReportNo = mail_Request.ReportNo });
            Report_Result baseResult = GetPDF(model);
            string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", baseResult.TempFileName) : string.Empty;
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = "Mockup Wash – Test Fail",
                To = mail_Request.To,
                CC = mail_Request.CC,
                //Body = mailBody,
                //alternateView = plainView,
                FileonServer = new List<string> { FileName },
                IsShowAIComment = true,
                AICommentType = "Mockup Wash Test",
                StyleID = model.StyleID,
                SeasonID = model.SeasonID,
                BrandID = model.BrandID,
            };

            _MailService = new MailToolsService();
            string comment = _MailService.GetAICommet(sendMail_Request);
            string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
            string mailBody = MailTools.DataTableChangeHtml(dt, comment, buyReadyDate, out AlternateView plainView);

            sendMail_Request.Body = mailBody;
            sendMail_Request.alternateView = plainView;

            return MailTools.SendMail(sendMail_Request);
        }
    }
}
