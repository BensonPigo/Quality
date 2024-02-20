using ADOHelper.Utility;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using DatabaseObject.ResultModel;
using Library;
using System.Web.UI.WebControls;
using static Ict.Win.Design.DateTimeConverter;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class HeatTransferWashService
    {
        public HeatTransferWashProvider _Provider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;
        private MailToolsService _MailService;
        public BaseResult Create(HeatTransferWash_ViewModel model, string MDivision, string userid, out string NewReportNo)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new HeatTransferWashProvider(_ISQLDataTransaction);
            NewReportNo = string.Empty;
            int ctn = 0;
            try
            {
                model.Details = model.Details == null ? new List<HeatTransferWash_Detail_Result>() : model.Details;

                //if (model.Details.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                //{
                //    model.Main.Result = "Fail";
                //}
                //else
                //{
                //    model.Main.Result = "Pass";
                //}

                ctn = _Provider.Insert_HeatTransferWash(model.Main, MDivision, userid, out NewReportNo);

                if (ctn == 0)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrorMessage = "Create data fail.";
                    return result;
                }
                model.Main.ReportNo = NewReportNo;

                foreach (var detail in model.Details)
                {
                    detail.ReportNo = NewReportNo;
                    detail.EditName = model.Main.AddName;
                    ctn = _Provider.Insert_HeatTransferWash_Detail(detail);
                    if (ctn == 0)
                    {
                        _ISQLDataTransaction.RollBack();
                        result.Result = false;
                        result.ErrorMessage = "Create detail Fail.";
                        return result;
                    }
                }

                _ISQLDataTransaction.Commit();
                result.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public BaseResult Update(HeatTransferWash_ViewModel model, string userid)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new HeatTransferWashProvider(_ISQLDataTransaction);
            int ctn = 0;
            try
            {
                model.Details = model.Details == null ? new List<HeatTransferWash_Detail_Result>() : model.Details;

                //if (model.Details.Any(a => a.Result.ToUpper() == "Fail".ToUpper()))
                //{
                //    model.Main.Result = "Fail";
                //}
                //else
                //{
                //    model.Main.Result = "Pass";
                //}

                model.Main.EditName = userid;
                ctn = _Provider.Update_HeatTransferWash(model);

                if (ctn == 0)
                {
                    _ISQLDataTransaction.RollBack();
                    result.Result = false;
                    result.ErrorMessage = "update data fail.";
                    return result;
                }

                foreach (var detail in model.Details)
                {
                    detail.ReportNo = model.Main.ReportNo;
                    detail.EditName = userid;
                }
                _Provider.Update_HeatTransferWash_Detail(model);

                _ISQLDataTransaction.Commit();
                result.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public BaseResult Delete(HeatTransferWash_ViewModel model)
        {
            BaseResult result = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);
            _Provider = new HeatTransferWashProvider(_ISQLDataTransaction);

            try
            {
                _Provider.Delete_HeatTransferWash(model.Request.ReportNo);

                _ISQLDataTransaction.Commit();
                result.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return result;
        }

        public BaseResult EncodeAmend(HeatTransferWash_ViewModel model)
        {
            BaseResult result = new BaseResult();

            try
            {
                _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);


                _Provider.Confirm_HeatTransferWash(model);

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }
            return result;
        }
        public HeatTransferWash_ViewModel GetHeatTransferWash(HeatTransferWash_Request Req)
        {
            HeatTransferWash_ViewModel model = new HeatTransferWash_ViewModel()
            {
                Main = new HeatTransferWash_Result(),
                Details = new List<HeatTransferWash_Detail_Result>(),
                Request = Req,
            };

            try
            {
                _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<SelectListItem> ReportNoList = _Provider.GetReportNo_Source(Req);

                if (!string.IsNullOrEmpty(Req.ReportNo))
                {
                    if (ReportNoList.Where(o => o.Value == Req.ReportNo).Any())
                    {
                        model.Main = _Provider.GetMainData(Req);
                        //model.Request = Req;
                        model.Details = _Provider.GetDetailData(Req.ReportNo).ToList();
                    }
                    else
                    {
                        if (ReportNoList.Any())
                        {
                            model.Main = _Provider.GetMainData(new HeatTransferWash_Request()
                            {
                                ReportNo = ReportNoList.FirstOrDefault().Value,
                            });
                            model.Details = _Provider.GetDetailData(ReportNoList.FirstOrDefault().Value).ToList();
                        }
                    }
                }
                else if (string.IsNullOrEmpty(Req.ReportNo) && ReportNoList.Any())
                {
                    model.Main = _Provider.GetMainData(new HeatTransferWash_Request()
                    {
                        ReportNo = ReportNoList.FirstOrDefault().Value,
                    });
                    model.Details = _Provider.GetDetailData(ReportNoList.FirstOrDefault().Value).ToList();
                }
                model.ReportNo_Source = ReportNoList;
                model.Request.ReportNo = model.Main.ReportNo;
                model.Request.BrandID = model.Main.BrandID;
                model.Request.SeasonID = model.Main.SeasonID;
                model.Request.StyleID = model.Main.StyleID;
                model.Request.Article = model.Main.Article;

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public BaseResult ToReport(string ReportNo ,bool IsPDF,  out string FinalFilenmae, string AssignedFineName = "")
        {

            BaseResult result = new BaseResult();
            _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);
            FinalFilenmae = string.Empty;

            try
            {
                HeatTransferWash_Result head = _Provider.GetMainData(new HeatTransferWash_Request() { ReportNo = ReportNo });
                List<HeatTransferWash_Detail_Result> body =_Provider.GetDetailData(ReportNo).ToList();

                System.Data.DataTable ReportTechnician = _Provider.GetReportTechnician(new HeatTransferWash_Request() { ReportNo = ReportNo });

                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string strXltName = baseFilePath + "\\XLT\\HeatTransferWash.xltx";
                Excel.Application excel = MyUtility.Excel.ConnectExcel(strXltName);
                if (excel == null)
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }
                excel.Visible = false;
                excel.DisplayAlerts = false;

                Excel.Worksheet worksheet= excel.ActiveWorkbook.Worksheets[1];


                // 表頭填入

                worksheet.Cells[1, 1] = head.ArtworkTypeID + " - Daily wash test report";

                worksheet.Cells[2, 2] = head.OrderID;
                worksheet.Cells[2, 7] = head.ReportDate.HasValue ? head.ReportDate.Value.ToString("yyyy-MM-dd") : string.Empty;

                worksheet.Cells[3, 2] = head.ReceivedDate.HasValue ? head.ReceivedDate.Value.ToString("yyyy-MM-dd") : string.Empty;
                worksheet.Cells[3, 7] = head.AddDate.HasValue ? head.AddDate.Value.ToString("yyyy-MM-dd") : string.Empty;

                worksheet.Cells[4, 2] = head.SeasonID;
                worksheet.Cells[4, 7] = head.Teamwear ? "V" : string.Empty;

                worksheet.Cells[5, 2] = head.Article;
                worksheet.Cells[5, 7] = head.StyleID;

                worksheet.Cells[6, 2] = head.BrandID;
                worksheet.Cells[6, 7] = head.Line;

                worksheet.Cells[7, 2] = head.ArtworkTypeID;
                worksheet.Cells[7, 7] = head.Machine;

                worksheet.Cells[8, 3] = head.ArtworkTypeID + " machine parameter";

                if (head.ArtworkTypeID == "BO" || head.ArtworkTypeID == "FU")
                {
                    worksheet.Cells[9, 2] = "Film ref#";
                }
                else if (head.ArtworkTypeID == "HT")
                {
                    worksheet.Cells[9, 2] = "HT ref#";
                }

                //worksheet.Cells[10, 3] = head.Time;
                //worksheet.Cells[11, 3] = head.Pressure;
                //worksheet.Cells[12, 3] = head.PeelOff;
                //worksheet.Cells[14, 1] = head.Cycles;
                //worksheet.Cells[14, 3] = head.TemperatureUnit;

                if (head.Result == "Pass")
                {
                    worksheet.Cells[12, 2] = "V";
                }
                else if (head.Result == "Fail")
                {
                    worksheet.Cells[12, 7] = "V";
                }
                worksheet.Cells[14, 1] = head.Remark;

                string imgPath_BeforePicture = string.Empty;
                string imgPath_AfterPicture = string.Empty;
                string imgPath_Signture = string.Empty;
                if (head.TestBeforePicture != null )
                {
                    byte[] beforePic = head.TestBeforePicture;
                    imgPath_BeforePicture = ToolKit.PublicClass.AddImageSignWord(beforePic, head.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    Excel.Range cell = worksheet.Cells[17, 1];
                    worksheet.Shapes.AddPicture(imgPath_BeforePicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 2, cell.Top + 2, 220, 130);
                }
                if (head.TestBeforePicture != null)
                {
                    byte[] beforePic = head.TestAfterPicture;
                    imgPath_AfterPicture = ToolKit.PublicClass.AddImageSignWord(beforePic, head.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    Excel.Range cell = worksheet.Cells[17, 6];
                    worksheet.Shapes.AddPicture(imgPath_AfterPicture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 2, cell.Top + 2, 220, 130);
                }
                if (head.Signature != null)
                {
                    byte[] SignturePic = head.Signature;
                    imgPath_Signture = ToolKit.PublicClass.AddImageSignWord(SignturePic, head.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    Excel.Range cell = worksheet.Cells[30, 7];
                    worksheet.Shapes.AddPicture(imgPath_Signture, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 40, 20);
                }

                // Technician 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["Technician"].ToString();

                    // 姓名
                    worksheet.Cells[20, 6] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[30, 7];
                    if (ReportTechnician.Rows[0]["TechnicianSignture"] != DBNull.Value)
                    {

                        byte[] TestBeforePicture = (byte[])ReportTechnician.Rows[0]["TechnicianSignture"]; // 圖片的 byte[]

                        MemoryStream ms = new MemoryStream(TestBeforePicture);
                        System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                        string imageName = $"{Guid.NewGuid()}.jpg";
                        string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                        img.Save(imgPath);
                        worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    }
                }

                // ISP20230055 簽名、名字一起顯示
                worksheet.Cells[29, 7] = head.LastEditText;

                // 表身筆數處理
                if (!body.Any())
                {
                    worksheet.get_Range("A10").EntireRow.Delete();
                }
                else
                {
                    int copyCount = body.Count - 1;

                    for (int i = 0; i < copyCount; i++)
                    {
                        Excel.Range paste1 = worksheet.get_Range($"A{i + 10}", Type.Missing);
                        Excel.Range copyRow = worksheet.get_Range("A10").EntireRow;
                        paste1.Insert(Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }
                }


                // 表身填入
                int bodyStart = 10;
                foreach (var item in body)
                {
                    worksheet.Cells[bodyStart, 1] = item.FabricRefNo;
                    worksheet.Cells[bodyStart, 2] = item.HTRefNo;
                    worksheet.Cells[bodyStart, 3] = item.Temperature;
                    worksheet.Cells[bodyStart, 4] = item.Time;
                    worksheet.Cells[bodyStart, 5] = item.SecondTime;
                    worksheet.Cells[bodyStart, 6] = item.Pressure;
                    worksheet.Cells[bodyStart, 7] = item.PeelOff;
                    worksheet.Cells[bodyStart, 8] = item.Cycles;
                    worksheet.Cells[bodyStart, 9] = item.TemperatureUnit;
                    worksheet.Cells[bodyStart, 10] = item.Remark;
                    bodyStart++;
                }

                string tmpName= $"Daily HT Wash Test-{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}";

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }

                string fileName = $"{tmpName}.xlsx";
                string fullExcelFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                string filePdfName = $"{tmpName}.pdf";
                string fullPdfFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filePdfName);

                Excel.Workbook workBook = excel.ActiveWorkbook;
                if (IsPDF)
                {
                    Microsoft.Office.Interop.Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;

                    workBook.ExportAsFixedFormat(targetType, fullPdfFileName);
                    Marshal.ReleaseComObject(workBook);
                    FinalFilenmae = filePdfName;
                }
                else
                {
                    workBook.SaveAs(fullExcelFileName);
                    //excel.ActiveWorkbook.SaveAs(excelPath);
                    FinalFilenmae = fileName;
                }

                workBook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workBook);
                Marshal.ReleaseComObject(excel);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
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

        /// <summary>
        /// Annotation取得，請參考 ISP20230789
        /// </summary>
        /// <param name="orders"></param>
        /// <returns></returns>
        public List<SelectListItem> GetArtworkTypeList(Orders orders)
        {
            List<SelectListItem> result = new List<SelectListItem>();
            _Provider = new HeatTransferWashProvider(Common.ProductionDataAccessLayer);
            try
            {
                List<string> list = new List<string>();

                list= _Provider.GetArtworkTypeOri(orders).ToList();

                foreach (var item in list)
                {
                    string[] Annotations = Regex.Replace(item, @"[\d]", string.Empty).Split('+'); // 剖析Annotation


                    foreach (var Annotation in Annotations)
                    {
                        if (!result.Any(o => o.Value == Annotation))
                        {

                            result.Add(new SelectListItem()
                            {
                                Text = Annotation,
                                Value = Annotation,
                            });
                        }
                    }
                }
                result = result.OrderBy(o => o.Value).ToList();
                return result;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public HeatTransferWash_Detail_Result GetLastDetailData(string HTRefNo)
        {
            HeatTransferWash_Detail_Result result = new HeatTransferWash_Detail_Result();
            try
            {
                _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);
                result =  _Provider.GetLastDetailData(HTRefNo);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public SendMail_Result SendMail(string ReportNo, string TO, string CC)
        {
            _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);

            HeatTransferWash_ViewModel model = this.GetHeatTransferWash(new HeatTransferWash_Request() { ReportNo = ReportNo });
            string FinalFilenmae = string.Empty;

            string name = $"Daily Heat Transfer Wash Test_{model.Main.OrderID}_" +
                        $"{model.Main.StyleID}_" +
                        $"{model.Main.Article}_" +
                        $"{model.Main.Line}_" +
                        $"{model.Main.Result}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            BaseResult report = this.ToReport(ReportNo, false ,out FinalFilenmae, name);

            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", FinalFilenmae);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Daily Heat Transfer Wash Test/{model.Main.OrderID}/" +
                        $"{model.Main.StyleID}/" +
                        $"{model.Main.Article}/" +
                        $"{model.Main.Line}/" +
                        $"{model.Main.Result}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                To = TO,
                CC = CC,
                Body = mailBody,
                //alternateView = plainView,
                FileonServer = new List<string> { FileName },
                IsShowAIComment = true,
            };

            _MailService = new MailToolsService();
            string comment = _MailService.GetAICommet(sendMail_Request);
            string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
            sendMail_Request.Body = sendMail_Request.Body + Environment.NewLine + comment + Environment.NewLine + buyReadyDate;

            return MailTools.SendMail(sendMail_Request);
        }
    }
}
