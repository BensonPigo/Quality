using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Office.Interop.Excel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Web.Mvc;
using System.Windows;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace BusinessLogicLayer.Service
{
    public class MockupOvenService : IMockupOvenService
    {
        private IMockupOvenProvider _MockupOvenProvider;
        private IMockupOvenDetailProvider _MockupOvenDetailProvider;
        private IStyleBOAProvider _IStyleBOAProvider;
        private IStyleArtworkProvider _IStyleArtworkProvider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;
        private IScaleProvider _ScaleProvider;
        private IInspectionTypeProvider _InspectionTypeProvider;
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
                _MockupOvenDetailProvider = new MockupOvenDetailProvider(Common.ProductionDataAccessLayer);
                mockupOven_model = _MockupOvenProvider.GetMockupOven(MockupOven, istop1: true).ToList().FirstOrDefault();
                if (mockupOven_model != null)
                {
                    mockupOven_model.ScaleID_Source = GetScale();
                    mockupOven_model.ReportNo_Source = _MockupOvenProvider.GetMockupOvenReportNoList(MockupOven).Select(s => s.ReportNo).ToList();
                    MockupOven_Detail mockupOven_Detail = new MockupOven_Detail() { ReportNo = mockupOven_model.ReportNo };
                    mockupOven_model.MockupOven_Detail = _MockupOvenDetailProvider.GetMockupOven_Detail(mockupOven_Detail).ToList();

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
                if (!(IsTest.ToLower() == "true"))
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

                tmpName = $"Mockup Oven _{mockupOven.POID}_" +
                    $"{mockupOven.StyleID}_" +
                    $"{mockupOven.Article}_" +
                    $"{mockupOven.Result}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";


                _MockupOvenProvider = new MockupOvenProvider(Common.ProductionDataAccessLayer);
                _InspectionTypeProvider = new InspectionTypeProvider(Common.ProductionDataAccessLayer);

                List<InspectionType> InspectionTypes = _InspectionTypeProvider.Get_InspectionType("MockupOven", "Bulk", mockupOven.BrandID).ToList();
                mockupOven.Requirements = InspectionTypes.Select(x => x.Comment).ToList();
                var mockupOven_Detail = mockupOven.MockupOven_Detail;

                bool haveHT = mockupOven.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER");
                string basefileName = haveHT ? "MockupOven2" : "MockupOven";
                string openfilepath;
                if (IsTest.ToLower() == "true")
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
                worksheet.Cells[4, 2] = mockupOven.ReportNo;
                worksheet.Cells[5, 2] = mockupOven.T1Subcon + "-" + mockupOven.T1SubconAbb;
                worksheet.Cells[6, 2] = mockupOven.T2Supplier + "-" + mockupOven.T2SupplierAbb;
                worksheet.Cells[7, 2] = mockupOven.BrandID;
                worksheet.Cells[8, 2] = $"5.14 color migration test({mockupOven.TestTemperature} degree @ {mockupOven.TestTime} hours)";

                worksheet.Cells[4, 8] = mockupOven.ReleasedDate;
                worksheet.Cells[5, 8] = mockupOven.TestDate;
                worksheet.Cells[6, 8] = mockupOven.SeasonID;

                if (haveHT)
                {
                    worksheet.Cells[10, 2] = mockupOven.HTPlate;
                    worksheet.Cells[11, 2] = mockupOven.HTFlim;
                    worksheet.Cells[12, 2] = mockupOven.HTTime;
                    worksheet.Cells[13, 2] = mockupOven.HTPressure;
                    worksheet.Cells[10, 8] = mockupOven.HTPellOff;
                    worksheet.Cells[11, 8] = mockupOven.HT2ndPressnoreverse;
                    worksheet.Cells[12, 8] = mockupOven.HT2ndPressreversed;
                    worksheet.Cells[13, 8] = mockupOven.HTCoolingTime;
                }

                worksheet.Cells[13 + haveHTrow, 2] = mockupOven.TechnicianName;

                Range cell = worksheet.Cells[12 + haveHTrow, 2];
                if (mockupOven.Signature != null)
                {
                    MemoryStream ms = new MemoryStream(mockupOven.Signature);
                    Image img = Image.FromStream(ms);
                    string imageName = $"{Guid.NewGuid()}.jpg";
                    string imgPath;
                    if (IsTest.ToLower() == "true")
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

                Range pic = worksheet.get_Range($"B{16 + haveHTrow}:E{34 + haveHTrow}");
                cell = worksheet.Cells[13 + haveHTrow + 3, 2];
                if (mockupOven.TestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(mockupOven.TestBeforePicture, mockupOven.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, pic.Left, pic.Top, pic.Width, pic.Height);
                }

                pic = worksheet.get_Range($"H{16 + haveHTrow}:N{34 + haveHTrow}");
                if (mockupOven.TestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(mockupOven.TestAfterPicture, mockupOven.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: IsTest.ToLower() == "true");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, pic.Left, pic.Top, pic.Width, pic.Height);
                }

                #region 表身資料
                // 插入多的row
                if (mockupOven_Detail.Count > 0)
                {
                    Range rngToInsert = worksheet.get_Range($"A{10 + haveHTrow}:N{10 + haveHTrow}", Type.Missing).EntireRow;
                    for (int i = 1; i < mockupOven_Detail.Count; i++)
                    {
                        rngToInsert.Insert(XlInsertShiftDirection.xlShiftDown);
                        worksheet.get_Range(string.Format("C{0}:D{0}", (10 + haveHTrow + i - 1).ToString())).Merge(false);
                        worksheet.get_Range(string.Format("I{0}:K{0}", (10 + haveHTrow + i - 1).ToString())).Merge(false);
                    }

                    worksheet.get_Range(string.Format("L{0}:N{1}", (10 + haveHTrow).ToString(), (10 + haveHTrow + mockupOven_Detail.Count - 1).ToString())).Merge(false);
                    Marshal.ReleaseComObject(rngToInsert);
                }

                // ISP20230792
                if ((mockupOven.HTPlate > 0)
                    || (mockupOven.HTFlim > 0)
                    || (mockupOven.HTTime > 0)
                    || (mockupOven.HTPressure > 0)
                    || !string.IsNullOrEmpty(mockupOven.HTPellOff) 
                    || (mockupOven.HT2ndPressnoreverse > 0)
                    || (mockupOven.HT2ndPressreversed > 0)
                    || (mockupOven.HTCoolingTime > 0))
                {
                    int aRow = 11 + haveHTrow + mockupOven_Detail.Count - 1;
                    Range rngToCopy = worksheet2.get_Range($"A1:N5", Type.Missing).EntireRow;
                    Range rngToInsert = worksheet.get_Range($"A{aRow}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                    rngToInsert.Insert(XlInsertShiftDirection.xlShiftDown, rngToCopy.Copy(Type.Missing)); // 貼上

                    worksheet.Cells[aRow + 1, 3] = mockupOven.HTPlate;
                    worksheet.Cells[aRow + 2, 3] = mockupOven.HTFlim;
                    worksheet.Cells[aRow + 3, 3] = mockupOven.HTTime;
                    worksheet.Cells[aRow + 4, 3] = mockupOven.HTPressure;
                    worksheet.Cells[aRow + 1, 11] = mockupOven.HTPellOff;
                    worksheet.Cells[aRow + 2, 11] = mockupOven.HT2ndPressnoreverse;
                    worksheet.Cells[aRow + 3, 11] = mockupOven.HT2ndPressreversed;
                    worksheet.Cells[aRow + 4, 11] = mockupOven.HTCoolingTime;
                }

                // 塞進資料
                int start_row = 10 + haveHTrow;
                worksheet.Cells[start_row, 12] = mockupOven.Requirements.JoinToString(Environment.NewLine);
                foreach (var item in mockupOven_Detail)
                {
                    string remark = item.Remark;
                    string fabric = string.IsNullOrEmpty(item.FabricColorName) ? item.FabricRefNo : item.FabricRefNo + " - " + item.FabricColorName;
                    string artwork = string.IsNullOrEmpty(mockupOven.ArtworkTypeID) ? item.Design + " - " + item.ArtworkColorName : mockupOven.ArtworkTypeID + "/" + item.Design + " - " + item.ArtworkColorName;
                    worksheet.Cells[start_row, 1] = mockupOven.StyleID;
                    worksheet.Cells[start_row, 2] = fabric;
                    worksheet.Cells[start_row, 3] = artwork;
                    worksheet.Cells[start_row, 5] = item.ChangeScale;
                    worksheet.Cells[start_row, 6] = item.ResultChange;
                    worksheet.Cells[start_row, 7] = item.StainingScale;
                    worksheet.Cells[start_row, 8] = item.ResultStain;
                    worksheet.Cells[start_row, 9] = item.Remark;                    
                    worksheet.Rows[start_row].Font.Bold = false;
                    worksheet.Rows[start_row].WrapText = true;
                    worksheet.Rows[start_row].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    int maxLength = fabric.Length > remark.Length ? fabric.Length : remark.Length;
                    maxLength = maxLength > artwork.Length ? maxLength : artwork.Length;
                    worksheet.Range[$"A{start_row}", $"N{start_row}"].RowHeight = ((maxLength / 20) + 1) * 16.5;

                    start_row++;
                }


                #endregion

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }

                string filexlsx = tmpName + ".xlsx";
                string fileNamePDF = tmpName + ".pdf";

                string filepath;
                string filepathpdf;
                if (IsTest.ToLower() == "true")
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

        public BaseResult Create(MockupOven_ViewModel MockupOven, string Mdivision, string userid, out string NewReportNo)
        {
            NewReportNo = string.Empty;
            MockupOven.Type = "B";
            MockupOven.AddName = userid;
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupOvenProvider = new MockupOvenProvider(_ISQLDataTransaction);
            _MockupOvenDetailProvider = new MockupOvenDetailProvider(_ISQLDataTransaction);
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

                count = _MockupOvenProvider.Create(MockupOven, Mdivision, out NewReportNo);
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
                    count = _MockupOvenDetailProvider.Create(MockupOven_Detail);
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
                _MockupOvenProvider.Update(MockupOven);
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
            _MockupOvenDetailProvider = new MockupOvenDetailProvider(_ISQLDataTransaction);
            try
            {
                _MockupOvenProvider.Delete(MockupOven);
                _MockupOvenDetailProvider.Delete(new MockupOven_Detail_ViewModel() { ReportNo = MockupOven.ReportNo });
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
            _MockupOvenDetailProvider = new MockupOvenDetailProvider(_ISQLDataTransaction);
            try
            {
                foreach (var MockupOven_Detail in MockupOvenDetail)
                {
                    MockupOven_Detail.ReportNo = null;
                    _MockupOvenDetailProvider.Delete(MockupOven_Detail);
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
            //string mailBody = MailTools.DataTableChangeHtml(dt,"","", out AlternateView plainView);
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
