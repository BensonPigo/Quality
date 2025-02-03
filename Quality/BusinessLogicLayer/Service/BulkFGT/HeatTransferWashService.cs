using ADOHelper.Utility;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using System.Text.RegularExpressions;
using DatabaseObject.ResultModel;
using Library;
using System.Web;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.EMMA;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class HeatTransferWashService
    {
        public HeatTransferWashProvider _Provider;
        private IOrdersProvider _OrdersProvider;
        private IOrderQtyProvider _OrderQtyProvider;
        private MailToolsService _MailService;
        QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;
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

                string Subject = $"Daily Heat Transfer Wash Test/{model.Main.OrderID}/" +
                        $"{model.Main.StyleID}/" +
                        $"{model.Main.Article}/" +
                        $"{model.Main.Result}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                model.Main.MailSubject = Subject;
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

        public BaseResult ToReport(string ReportNo, bool IsPDF, out string FinalFilenmae, string AssignedFineName = "")
        {
            BaseResult result = new BaseResult();
            _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
            FinalFilenmae = string.Empty;
            string tmpName = string.Empty;

            try
            {
                HeatTransferWash_Result head = _Provider.GetMainData(new HeatTransferWash_Request() { ReportNo = ReportNo });
                var testCode = _QualityBrandTestCodeProvider.Get(head.BrandID, "Daily HT Wash");
                List<HeatTransferWash_Detail_Result> body = _Provider.GetDetailData(ReportNo).ToList();

                System.Data.DataTable ReportTechnician = _Provider.GetReportTechnician(new HeatTransferWash_Request() { ReportNo = ReportNo });

                tmpName = $"Daily Heat Transfer Wash Test_{head.OrderID}_" +
                       $"{head.StyleID}_" +
                       $"{head.Article}_" +
                       $"{head.Result}_" +
                       $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                string strXltName = baseFilePath + $@"\XLT\HeatTransferWash.xltx";

                if (!File.Exists(strXltName))
                {
                    result.ErrorMessage = "Excel template not found!";
                    result.Result = false;
                    return result;
                }

                using (var workbook = new XLWorkbook(strXltName))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 表頭填入
                    if (testCode.Any())
                    {
                        worksheet.Cell(1, 1).Value = head.ArtworkTypeID + $" - Daily wash test report({testCode.FirstOrDefault().TestCode})";
                    }
                    else
                    {
                        worksheet.Cell(1, 1).Value = head.ArtworkTypeID + " - Daily wash test report";
                    }

                    worksheet.Cell(2, 2).Value = head.OrderID;
                    worksheet.Cell(2, 7).Value = head.ReportDate?.ToString("yyyy-MM-dd") ?? string.Empty;

                    worksheet.Cell(3, 2).Value = head.ReceivedDate?.ToString("yyyy-MM-dd") ?? string.Empty;
                    worksheet.Cell(3, 7).Value = head.AddDate?.ToString("yyyy-MM-dd") ?? string.Empty;

                    worksheet.Cell(4, 2).Value = head.SeasonID;
                    worksheet.Cell(4, 7).Value = head.Teamwear ? "V" : string.Empty;

                    worksheet.Cell(5, 2).Value = head.Article;
                    worksheet.Cell(5, 7).Value = head.StyleID;

                    worksheet.Cell(6, 2).Value = head.BrandID;
                    worksheet.Cell(6, 7).Value = head.Line;

                    worksheet.Cell(7, 2).Value = head.ArtworkTypeID;
                    worksheet.Cell(7, 7).Value = head.Machine;

                    worksheet.Cell(8, 3).Value = head.ArtworkTypeID + " machine parameter";

                    if (head.ArtworkTypeID == "BO" || head.ArtworkTypeID == "FU")
                    {
                        worksheet.Cell(9, 2).Value = "Film ref#";
                    }
                    else if (head.ArtworkTypeID == "HT")
                    {
                        worksheet.Cell(9, 2).Value = "HT ref#";
                    }

                    if (head.Result == "Pass")
                    {
                        worksheet.Cell(12, 2).Value = "V";
                    }
                    else if (head.Result == "Fail")
                    {
                        worksheet.Cell(12, 7).Value = "V";
                    }

                    worksheet.Cell(14, 1).Value = head.Remark;

                    // 插入圖片
                    AddImageToWorksheet(worksheet, head.TestBeforePicture, 17, 1, 220, 130);
                    AddImageToWorksheet(worksheet, head.TestAfterPicture, 17, 6, 220, 130);
                    AddImageToWorksheet(worksheet, head.Signature, 30, 7, 40, 20);

                    // 表身筆數處理，複製儲存格
                    if (!body.Any())
                    {
                        worksheet.Row(10).Delete();

                    }
                    else
                    {
                        int copyCount = body.Count - 1;

                        for (int i = 0; i < copyCount; i++)
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


                    // 表身填入
                    int bodyStart = 10;
                    foreach (var item in body)
                    {
                        worksheet.Cell(bodyStart, 1).Value = item.FabricRefNo;
                        worksheet.Cell(bodyStart, 2).Value = item.HTRefNo;
                        worksheet.Cell(bodyStart, 3).Value = item.Temperature;
                        worksheet.Cell(bodyStart, 4).Value = item.Time;
                        worksheet.Cell(bodyStart, 5).Value = item.SecondTime;
                        worksheet.Cell(bodyStart, 6).Value = item.Pressure;
                        worksheet.Cell(bodyStart, 7).Value = item.PeelOff;
                        worksheet.Cell(bodyStart, 8).Value = item.Cycles;
                        worksheet.Cell(bodyStart, 9).Value = item.TemperatureUnit;
                        worksheet.Cell(bodyStart, 10).Value = item.Remark;
                        bodyStart++;
                    }

                    #region Title
                    string FactoryNameEN = _Provider.GetFactoryNameEN(ReportNo, System.Web.HttpContext.Current.Session["FactoryID"].ToString());
                    // 1. 插入一列
                    worksheet.Row(1).InsertRowsAbove(1);

                    // 2. 合併欄位
                    worksheet.Range("A1:J1").Merge();
                    // 設置字體樣式
                    var mergedCell = worksheet.Cell("A1");
                    mergedCell.Value = FactoryNameEN;
                    mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                    mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                    mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    mergedCell.Style.Font.Bold = true;
                    mergedCell.Style.Font.Italic = false;

                    // 自動檢測使用範圍
                    var usedRange = worksheet.RangeUsed();
                    var lastRow = worksheet.CellsUsed().Max(cell => cell.Address.RowNumber);
                    // 確認範圍不為空
                    if (usedRange != null)
                    {
                        // 清除所有已有的列印範圍
                        worksheet.PageSetup.PrintAreas.Clear();

                        // 設定列印範圍為使用範圍
                        worksheet.PageSetup.PrintAreas.Add($"A1:J{lastRow + 10}");
                    }
                    #endregion

                    // 儲存檔案
                    if (!string.IsNullOrWhiteSpace(AssignedFineName))
                    {
                        tmpName = AssignedFineName;
                    }

                    tmpName = RemoveInvalidFileNameChars(tmpName);

                    string fileName = $"{tmpName}.xlsx";
                    string fullExcelFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                    workbook.SaveAs(fullExcelFileName);


                    if (IsPDF)
                    {
                        LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                        officeService.ConvertExcelToPdf(fullExcelFileName, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                        FinalFilenmae = $"{tmpName}.pdf";
                    }
                    else
                    {
                        FinalFilenmae = fileName;
                    }
                    result.Result = true;
                }
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
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

                list = _Provider.GetArtworkTypeOri(orders).ToList();

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
                result = _Provider.GetLastDetailData(HTRefNo);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new HeatTransferWashProvider(Common.ManufacturingExecutionDataAccessLayer);

            HeatTransferWash_ViewModel model = this.GetHeatTransferWash(new HeatTransferWash_Request() { ReportNo = ReportNo });
            string FinalFilenmae = string.Empty;

            string name = $"Daily Heat Transfer Wash Test_{model.Main.OrderID}_" +
                        $"{model.Main.StyleID}_" +
                        $"{model.Main.Article}_" +
                        $"{model.Main.Result}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            BaseResult report = this.ToReport(ReportNo, false, out FinalFilenmae, name);

            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", FinalFilenmae);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Daily Heat Transfer Wash Test/{model.Main.OrderID}/" +
                        $"{model.Main.StyleID}/" +
                        $"{model.Main.Article}/" +
                        $"{model.Main.Result}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                To = TO,
                CC = CC,
                Body = mailBody,
                //alternateView = plainView,
                FileonServer = new List<string> { FileName },
                FileUploader = Files,
                IsShowAIComment = true,
            };

            if (!string.IsNullOrEmpty(Subject))
            {
                sendMail_Request.Subject = Subject;
            }

            _MailService = new MailToolsService();
            string comment = _MailService.GetAICommet(sendMail_Request);
            string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);

            sendMail_Request.Body = Body + Environment.NewLine + sendMail_Request.Body + Environment.NewLine + comment + Environment.NewLine + buyReadyDate;

            return MailTools.SendMail(sendMail_Request);
        }
    }
}
