using ADOHelper.Utility;
using DatabaseObject;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ProductionDataAccessLayer.Provider.MSSQL.BukkFGT;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class AccessoryOvenWashService
    {
        private AccessoryOvenWashProvider _AccessoryOvenWashProvider;
        private string IsTest = ConfigurationManager.AppSettings["IsTest"].ToString();

        public Accessory_ViewModel GetMainData(Accessory_ViewModel Req)
        {
            Accessory_ViewModel result = new Accessory_ViewModel()
            {
                ReqOrderID = Req.ReqOrderID,
                DataList = new List<Accessory_Result>(),
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetHead(Req.ReqOrderID);
                result.DataList = _AccessoryOvenWashProvider.GetDetail(Req.ReqOrderID).ToList();

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }

            return result;
        }

        public Accessory_ViewModel Update(Accessory_ViewModel Req)
        {
            Accessory_ViewModel result = new Accessory_ViewModel()
            {
                ReqOrderID = Req.ReqOrderID,
                DataList = new List<Accessory_Result>(),
            };
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);

                int r = _AccessoryOvenWashProvider.Update_AIR_Laboratory(Req);

                result.Result = r > 0;

                _AccessoryOvenWashProvider.Update_AIR_Laboratory_AllResult(Req);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public void UpdateInspPercent(string POID)
        {
            _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
            _AccessoryOvenWashProvider.UpdateInspPercent(POID);
        }

        #region Oven
        public Accessory_Oven GetOvenTest(Accessory_Oven Req)
        {

            Accessory_Oven result = new Accessory_Oven()
            {
                ErrorMessage = Req.ErrorMessage,
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetOvenTest(Req);
                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }

            return result;
        }

        public Accessory_Oven UpdateOven(Accessory_Oven Req)
        {
            Accessory_Oven result = new Accessory_Oven();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                int r = _AccessoryOvenWashProvider.UpdateOvenTest(Req);

                result.Result = r > 0;
                _AccessoryOvenWashProvider.Update_Oven_AllResult(Req);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return result;
        }

        public SendMail_Result SendOvenMail(Accessory_Oven Req)
        {

            SendMail_Result result = new SendMail_Result();
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
                DataTable dt = _AccessoryOvenWashProvider.GetData_OvenDataTable(Req);
                string mailBody = MailTools.DataTableChangeHtml(dt);

                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = Req.ToAddress,
                    CC = Req.CcAddress,
                    Subject = "Accessory Oven Test - Test Fail",
                    Body = mailBody
                };
                result = MailTools.SendMail(sendMail_Request);
                result.result = true;
            }
            catch (Exception ex)
            {

                result.result = false;
                result.resultMsg = ex.Message.ToString();
            }


            return result;
        }


        public BaseResult OvenTestExcel(string AIR_LaboratoryID, string POID, string Seq1, string Seq2 ,bool isPDF, out string FileName)
        {
            _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

            BaseResult result = new BaseResult();

            Accessory_OvenExcel Model = new Accessory_OvenExcel();

            FileName = string.Empty;

            try
            {
                string baseFilePath =  System.Web.HttpContext.Current.Server.MapPath("~/");
                //DataTable dtOvenDetail = _AccessoryOvenWashProvider.GetOvenTestDataTable(new Accessory_Oven() 
                //{ 
                //    AIR_LaboratoryID = Convert.ToInt64( AIR_LaboratoryID),
                //    POID = POID,
                //    Seq1 = Seq1,
                //    Seq2 = Seq2,
                //});

                Model =  _AccessoryOvenWashProvider.GetOvenTestExcel(new Accessory_Oven()
                {
                    AIR_LaboratoryID = Convert.ToInt64(AIR_LaboratoryID),
                    POID = POID,
                    Seq1 = Seq1,
                    Seq2 = Seq2,
                });

                if (Model == null)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                string strXltName = baseFilePath + "\\XLT\\AccessoryOvenTest.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                excel.DisplayAlerts = false;
                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];

                worksheet.Cells[2, 6] = Model.POID;
                worksheet.Cells[2, 10] = Model.Supplier;

                worksheet.Cells[3, 2] = Model.BrandID;
                worksheet.Cells[3, 6] = Model.Refno;
                worksheet.Cells[3, 10] = Model.WKNo;

                worksheet.Cells[4, 2] = "Bulk";
                worksheet.Cells[4, 6] = Model.Color;
                worksheet.Cells[4, 10] = string.Empty;

                worksheet.Cells[5, 2] = Model.StyleID;
                worksheet.Cells[5, 6] = Model.Size;
                worksheet.Cells[5, 10] =Model.SeasonID;
                worksheet.Cells[5, 12] = Model.Seq;

                worksheet.Cells[9, 2] = Model.OvenDate.HasValue ? Model.OvenDate.Value.ToString("yyyy/MM/dd") : string.Empty;
                worksheet.Cells[9, 10] = DateTime.Now.ToString("yyyy/MM/dd");

                worksheet.Cells[12, 2] = Model.OvenResult;
                worksheet.Cells[12, 5] = string.Empty;
                worksheet.Cells[12, 10] = string.Empty;

                worksheet.Cells[13, 2] = Model.Remark;

                worksheet.Cells[22, 3] = Model.OvenInspector;
                worksheet.Cells[22, 9] = Model.OvenInspector;

                #region 添加圖片
                Excel.Range cellBeforePicture = worksheet.Cells[20, 1];
                if (Model.OvenTestBeforePicture != null)
                {
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

                    byte[] bytes = Model.OvenTestBeforePicture;
                    using (var imageFile = new FileStream(imgPath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 300, 300);
                }

                Excel.Range cellAfterPicture = worksheet.Cells[20, 7];
                if (Model.OvenTestAfterPicture != null)
                {
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

                    byte[] bytes = Model.OvenTestAfterPicture;
                    using (var imageFile = new FileStream(imgPath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 2, 300, 300);
                }

                #endregion


                #region Save & Show Excel

                string pdfFileName = $"AccessoryOvenTest{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
                FileName = $"AccessoryOvenTest{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";

                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", FileName);

                excel.ActiveWorkbook.SaveAs(excelPath);
                excel.Quit();

                if (isPDF)
                {
                    bool isCreatePdfOK = ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
                    FileName = pdfFileName;
                    if (!isCreatePdfOK)
                    {
                        result.Result = false;
                        result.ErrorMessage = "ConvertToPDF fail";
                        return result;
                    }
                }

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excel);
                #endregion
            }
            catch (Exception ex)
            {

                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }



            return result;
        }
        #endregion


        #region Wash
        public Accessory_Wash GetWashTest(Accessory_Wash Req)
        {

            Accessory_Wash result = new Accessory_Wash()
            {
                ErrorMessage = Req.ErrorMessage,
            };

            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

                result = _AccessoryOvenWashProvider.GetWashTest(Req);
                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }

            return result;
        }

        public Accessory_Wash UpdateWash(Accessory_Wash Req)
        {
            Accessory_Wash result = new Accessory_Wash();

            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(_ISQLDataTransaction);

                result.ScaleData = _AccessoryOvenWashProvider.GetScaleData();
                int r = _AccessoryOvenWashProvider.UpdateWashTest(Req);

                result.Result = r > 0;
                _AccessoryOvenWashProvider.Update_Wash_AllResult(Req);

                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = $@"msg.WithError(""{ex.Message}"");";
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public SendMail_Result SendWashMail(Accessory_Wash Req)
        {

            SendMail_Result result = new SendMail_Result();
            try
            {
                _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);
                DataTable dt = _AccessoryOvenWashProvider.GetData_WashDataTable(Req);
                string mailBody = MailTools.DataTableChangeHtml(dt);

                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = Req.ToAddress,
                    CC = Req.CcAddress,
                    Subject = "Accessory Wash Test - Test Fail",
                    Body = mailBody
                };
                result = MailTools.SendMail(sendMail_Request);
                result.result = true;
            }
            catch (Exception ex)
            {

                result.result = false;
                result.resultMsg = ex.Message.ToString();
            }


            return result;
        }

        public BaseResult WashTestExcel(string AIR_LaboratoryID, string POID, string Seq1, string Seq2, bool isPDF, out string FileName)
        {
            _AccessoryOvenWashProvider = new AccessoryOvenWashProvider(Common.ProductionDataAccessLayer);

            BaseResult result = new BaseResult();

            Accessory_WashExcel Model = new Accessory_WashExcel();

            FileName = string.Empty;

            try
            {
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                //DataTable dtOvenDetail = _AccessoryOvenWashProvider.GetOvenTestDataTable(new Accessory_Oven() 
                //{ 
                //    AIR_LaboratoryID = Convert.ToInt64( AIR_LaboratoryID),
                //    POID = POID,
                //    Seq1 = Seq1,
                //    Seq2 = Seq2,
                //});

                Model = _AccessoryOvenWashProvider.GetWashTestExcel(new Accessory_Wash()
                {
                    AIR_LaboratoryID = Convert.ToInt64(AIR_LaboratoryID),
                    POID = POID,
                    Seq1 = Seq1,
                    Seq2 = Seq2,
                });

                if (Model == null)
                {
                    result.ErrorMessage = "Data not found!";
                    result.Result = false;
                    return result;
                }

                string strXltName = baseFilePath + "\\XLT\\AccessoryWashTest.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                excel.DisplayAlerts = false;
                Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1];

                worksheet.Cells[2, 6] = Model.POID;
                worksheet.Cells[2, 10] = Model.Supplier;

                worksheet.Cells[3, 2] = Model.BrandID;
                worksheet.Cells[3, 6] = Model.Refno;
                worksheet.Cells[3, 10] = Model.WKNo;

                worksheet.Cells[4, 2] = "Bulk";
                worksheet.Cells[4, 6] = Model.Color;
                worksheet.Cells[4, 10] = string.Empty;

                worksheet.Cells[5, 2] = Model.StyleID;
                worksheet.Cells[5, 6] = Model.Size;
                worksheet.Cells[5, 10] = Model.SeasonID;
                worksheet.Cells[5, 12] = Model.Seq;

                worksheet.Cells[9, 2] = Model.WashDate.HasValue ? Model.WashDate.Value.ToString("yyyy/MM/dd") : string.Empty;
                worksheet.Cells[9, 10] = DateTime.Now.ToString("yyyy/MM/dd");

                worksheet.Cells[15, 2] = Model.WashResult;

                worksheet.Cells[17, 2] = Model.Remark;

                worksheet.Cells[22, 3] = Model.WashInspector;
                worksheet.Cells[22, 9] = Model.WashInspector;

                #region 添加圖片
                Excel.Range cellBeforePicture = worksheet.Cells[24, 1];
                if (Model.WashTestBeforePicture != null)
                {
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

                    byte[] bytes = Model.WashTestBeforePicture;
                    using (var imageFile = new FileStream(imgPath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 300, 300);
                }

                Excel.Range cellAfterPicture = worksheet.Cells[24, 7];
                if (Model.WashTestAfterPicture != null)
                {
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

                    byte[] bytes = Model.WashTestAfterPicture;
                    using (var imageFile = new FileStream(imgPath, FileMode.Create))
                    {
                        imageFile.Write(bytes, 0, bytes.Length);
                        imageFile.Flush();
                    }
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 2, 300, 300);
                }

                #endregion


                #region Save & Show Excel

                string pdfFileName = $"AccessoryWashTest{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
                FileName = $"AccessoryWashTest{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";

                string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
                string excelPath = Path.Combine(baseFilePath, "TMP", FileName);

                excel.ActiveWorkbook.SaveAs(excelPath);
                excel.Quit();

                if (isPDF)
                {
                    bool isCreatePdfOK = ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
                    FileName = pdfFileName;
                    if (!isCreatePdfOK)
                    {
                        result.Result = false;
                        result.ErrorMessage = "ConvertToPDF fail";
                        return result;
                    }
                }

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excel);
                #endregion
            }
            catch (Exception ex)
            {

                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }



            return result;
        }
        #endregion
    }
}
