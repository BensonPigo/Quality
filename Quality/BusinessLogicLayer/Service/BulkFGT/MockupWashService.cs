using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
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
    public class MockupWashService : IMockupWashService
    {
        private IMockupWashProvider _MockupWashProvider;
        private IMockupWashDetailProvider _MockupWashDetailProvider;
        private IStyleBOAProvider _IStyleBOAProvider;
        private IStyleArtworkProvider _IStyleArtworkProvider;
        private IDropDownListProvider _DropDownListProvider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;

        public MockupWash_ViewModel GetMockupWash(MockupWash_Request MockupWash)
        {
            MockupWash.Type = "B";
            MockupWash_ViewModel mockupWash_model = new MockupWash_ViewModel();
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
                throw ex;
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

        public MockupWash_ViewModel GetPDF(MockupWash_ViewModel mockupWash, bool test = false)
        {
            MockupWash_ViewModel result = new MockupWash_ViewModel();
            if (mockupWash == null)
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

                _MockupWashProvider = new MockupWashProvider(Common.ProductionDataAccessLayer);


                var mockupWash_Detail = mockupWash.MockupWash_Detail;
                string basefileName = "MockupWash";
                string openfilepath;
                if (test)
                {
                    openfilepath = "C:\\Git\\Quality\\Quality\\Quality\\bin\\XLT\\MockupWash.xltx";
                }
                else
                {
                    openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";
                }

                Application excelApp = MyUtility.Excel.ConnectExcel(openfilepath);



                excelApp.DisplayAlerts = false;
                Worksheet worksheet = excelApp.Sheets[1];

                bool haveHT = mockupWash.ArtworkTypeID.ToUpper().EqualString("HEAT TRANSFER");
                int htRow = 6;
                int haveHTrow = haveHT ? htRow : 0;

                // 設定表頭資料
                worksheet.Cells[4, 2] = mockupWash.ReportNo;
                worksheet.Cells[5, 2] = mockupWash.T1Subcon + "-" + mockupWash.T1SubconAbb; ;
                worksheet.Cells[6, 2] = mockupWash.T2Supplier + "-" + mockupWash.T2SupplierAbb; ;

                worksheet.Cells[7, 2] = mockupWash.TestingMethod;
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
                else
                {
                    for (int i = 0; i < htRow; i++)
                    {
                        worksheet.Rows[9].Delete(XlDeleteShiftDirection.xlShiftUp);
                    }
                }

                worksheet.Cells[13 + haveHTrow, 2] = mockupWash.TechnicianName;

                Range cell = worksheet.Cells[12 + haveHTrow, 2];

                string picSource = mockupWash.SignaturePic;
                if (!MyUtility.Check.Empty(picSource))
                {
                    if (File.Exists(picSource))
                    {
                        worksheet.Shapes.AddPicture(picSource, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);
                    }
                }

                #region 表身資料
                // 插入多的row
                if (mockupWash_Detail.Count > 0)
                {
                    Range rngToInsert = worksheet.get_Range($"A{10 + haveHTrow}:G{10 + haveHTrow}", Type.Missing).EntireRow;
                    for (int i = 1; i < mockupWash_Detail.Count; i++)
                    {
                        rngToInsert.Insert(XlInsertShiftDirection.xlShiftDown);
                        worksheet.get_Range(string.Format("E{0}:G{0}", (10 + haveHTrow + i - 1).ToString())).Merge(false);
                    }

                    Marshal.ReleaseComObject(rngToInsert);
                }

                // 塞進資料
                int start_row = 10 + haveHTrow;
                foreach (var item in mockupWash_Detail)
                {
                    string remark = item.Remark;
                    string fabric = string.IsNullOrEmpty(item.FabricColorName) ? item.FabricRefNo : item.FabricRefNo + " - " + item.FabricColorName;
                    string artwork = mockupWash.ArtworkTypeID + "/" + item.Design + " - " + item.ArtworkColor;
                    worksheet.Cells[start_row, 1] = mockupWash.StyleID;
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

        public BaseResult Create(MockupWash_ViewModel MockupWash)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupWashProvider = new MockupWashProvider(_ISQLDataTransaction);
            _MockupWashDetailProvider = new MockupWashDetailProvider(_ISQLDataTransaction);
            int count;
            try
            {
                count = _MockupWashProvider.Create(MockupWash);
                if (count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Create MockupWash Fail. 0 Count";
                    return result;
                }

                foreach (var MockupWash_Detail in MockupWash.MockupWash_Detail)
                {
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
                throw ex;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult Update(MockupWash_ViewModel MockupWash)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupWashProvider = new MockupWashProvider(_ISQLDataTransaction);
            _MockupWashDetailProvider = new MockupWashDetailProvider(_ISQLDataTransaction);
            int count;
            try
            {
                count = _MockupWashProvider.Update(MockupWash);
                if (count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Update MockupWash Fail. 0 Count";
                    return result;
                }

                foreach (var MockupWash_Detail in MockupWash.MockupWash_Detail)
                {
                    count = _MockupWashDetailProvider.Update(MockupWash_Detail);
                    if (count == 0)
                    {
                        result.Result = false;
                        result.ErrorMessage = "Update MockupWash_Detail Fail. 0 Count";
                        return result;
                    }
                }

                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Update MockupWash Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
                throw ex;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

        public BaseResult Delete(MockupWash_ViewModel MockupWash)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            _MockupWashProvider = new MockupWashProvider(_ISQLDataTransaction);
            int count;
            try
            {
                count = _MockupWashProvider.Delete(MockupWash);
                result.Result = true;
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = "Delete MockupWash Fail";
                result.Exception = ex;
                _ISQLDataTransaction.RollBack();
                throw ex;
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
                throw ex;
            }
            finally { _ISQLDataTransaction.CloseConnection(); }
            return result;
        }

    }
}
