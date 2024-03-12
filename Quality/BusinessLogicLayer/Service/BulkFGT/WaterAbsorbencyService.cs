using ADOHelper.Utility;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.UI.WebControls;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class WaterAbsorbencyService
    {
        private WaterAbsorbencyProvider _Provider;
        private MailToolsService _MailService;
        public WaterAbsorbency_ViewModel GetDefaultModel(bool iNew = false)
        {
            WaterAbsorbency_ViewModel model = new WaterAbsorbency_ViewModel()
            {
                Request = new WaterAbsorbency_Request(),
                Main = new WaterAbsorbency_Main(),
                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<WaterAbsorbency_Detail>()
                {
                    new WaterAbsorbency_Detail() { EvaluationItem = "Before Wash" },
                    new WaterAbsorbency_Detail() { EvaluationItem = "After 5 Cycle Wash" },
                },
            };

            try
            {
                _Provider = new WaterAbsorbencyProvider(Common.ProductionDataAccessLayer);
                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }
            return model;
        }

        public WaterAbsorbency_ViewModel GetData(WaterAbsorbency_Request Req)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new WaterAbsorbencyProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<WaterAbsorbency_Main> tmpList = new List<WaterAbsorbency_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new WaterAbsorbency_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new WaterAbsorbency_Request()
                    {
                        BrandID = Req.BrandID,
                        SeasonID = Req.SeasonID,
                        StyleID = Req.StyleID,
                        Article = Req.Article,
                    });
                }


                if (tmpList.Any())
                {
                    // 取得 ReportNo下拉選單
                    foreach (var item in tmpList)
                    {
                        model.ReportNo_Source.Add(new System.Web.Mvc.SelectListItem()
                        {
                            Text = item.ReportNo,
                            Value = item.ReportNo,
                        });
                    }

                    // 取得表身
                    // 若"有"傳入ReportNo，則可以直接找出表頭表身明細
                    if (!string.IsNullOrEmpty(Req.ReportNo))
                    {
                        // 取得表頭資料
                        model.Main = tmpList.Where(o => o.ReportNo == Req.ReportNo).FirstOrDefault();
                    }
                    // 若"沒"傳入ReportNo，抓下拉選單第一筆顯示
                    else
                    {
                        // 根據下拉選單第一筆，取得表頭資料
                        model.Main = tmpList.Where(o => o.ReportNo == model.ReportNo_Source.FirstOrDefault().Value).FirstOrDefault();

                    }

                    // 取得Article 下拉選單
                    _Provider = new WaterAbsorbencyProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new WaterAbsorbency_Request() { OrderID = model.Main.OrderID });

                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    _Provider = new WaterAbsorbencyProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new WaterAbsorbency_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    string Subject = $"Water Absorbency Test/{model.Main.OrderID}/" +
                        $"{model.Main.StyleID}/" +
                        $"{model.Main.FabricRefNo}/" +
                        $"{model.Main.FabricColor}/" +
                        $"{model.Main.Result}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                    model.Main.MailSubject = Subject;
                    model.Result = true;
                }
                else
                {
                    model.Result = false;
                    model.ErrorMessage = "Data not found.";
                }

            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public WaterAbsorbency_ViewModel GetOrderInfo(string OrderID)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new WaterAbsorbencyProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new WaterAbsorbency_Request() { OrderID = OrderID });


                // 確認SP#是否存在
                if (tmpOrders.Any())
                {
                    // 取得表頭SP#相關欄位
                    model.Main.OrderID = tmpOrders.FirstOrDefault().ID;
                    model.Main.FactoryID = tmpOrders.FirstOrDefault().FactoryID;
                    model.Main.BrandID = tmpOrders.FirstOrDefault().BrandID;
                    model.Main.SeasonID = tmpOrders.FirstOrDefault().SeasonID;
                    model.Main.StyleID = tmpOrders.FirstOrDefault().StyleID;
                    model.Main.Article = tmpOrders.FirstOrDefault().Article;
                    model.Main.FabricType = tmpOrders.FirstOrDefault().FabricType;

                    model.Result = true;
                }
                else
                {
                    model.Result = false;
                    model.ErrorMessage = "SP# not found.";
                }

            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public WaterAbsorbency_ViewModel NewSave(WaterAbsorbency_ViewModel Req, string MDivision, string UserID)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WaterAbsorbencyProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<WaterAbsorbency_Detail>();
                }


                // 新增，並取得ReportNo
                _Provider.Insert_WaterAbsorbency(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // WaterAbsorbency_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<WaterAbsorbency_Detail>();
                }

                _Provider.Processe_WaterAbsorbency_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WaterAbsorbency_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                    ReportNo = Req.Main.ReportNo,
                });

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
                _ISQLDataTransaction.RollBack();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }
        public WaterAbsorbency_ViewModel EditSave(WaterAbsorbency_ViewModel Req, string UserID)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WaterAbsorbencyProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<WaterAbsorbency_Detail>();
                }


                // 更新表頭，並取得ReportNo
                _Provider.Update_WaterAbsorbency(Req, UserID);

                // 更新表身
                _Provider.Processe_WaterAbsorbency_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WaterAbsorbency_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                    ReportNo = Req.Main.ReportNo,
                });

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
                _ISQLDataTransaction.RollBack();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }
        public WaterAbsorbency_ViewModel Delete(WaterAbsorbency_ViewModel Req)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WaterAbsorbencyProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_WaterAbsorbency(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WaterAbsorbency_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new WaterAbsorbency_Request()
                {
                    ReportNo = model.Main.ReportNo,
                    BrandID = model.Main.BrandID,
                    SeasonID = model.Main.SeasonID,
                    StyleID = model.Main.StyleID,
                    Article = model.Main.Article,
                };

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
                _ISQLDataTransaction.RollBack();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }
        public WaterAbsorbency_ViewModel EncodeAmend(WaterAbsorbency_ViewModel Req, string UserID)
        {
            WaterAbsorbency_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WaterAbsorbencyProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_WaterAbsorbency(Req.Main, UserID);

                _ISQLDataTransaction.Commit();

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message.ToString();
                _ISQLDataTransaction.RollBack();
            }
            finally
            {
                _ISQLDataTransaction.CloseConnection();
            }

            return model;
        }

        public WaterAbsorbency_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            WaterAbsorbency_ViewModel result = new WaterAbsorbency_ViewModel();
            string basefileName = "WaterAbsorbencyTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";
            string tmpName = string.Empty;

            Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);

            try
            {
                _Provider = new WaterAbsorbencyProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得報表資料

                WaterAbsorbency_ViewModel model = this.GetData(new WaterAbsorbency_Request() { ReportNo = ReportNo });

                DataTable ReportTechnician = _Provider.GetReportTechnician(new WaterAbsorbency_Request() { ReportNo = ReportNo });

                tmpName = $"Water Absorbency Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";


                excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表

                // 取得工作表上所有圖形物件
                Microsoft.Office.Interop.Excel.Shapes shapes = worksheet.Shapes;

                // 根據名稱，搜尋文字方塊物件
                Microsoft.Office.Interop.Excel.Shape ADIDAS_TextBox = shapes.Item("ADIDAS_TextBox");
                Microsoft.Office.Interop.Excel.Shape REEBOK_TextBox = shapes.Item("REEBOK_TextBox");


                // BrandID
                if (model.Main.BrandID.ToUpper() == "ADIDAS")
                {
                    ADIDAS_TextBox.TextFrame.Characters().Text = "V";
                }
                if (model.Main.BrandID.ToUpper() == "REEBOK")
                {
                    REEBOK_TextBox.TextFrame.Characters().Text = "V";
                }

                string reportNo = model.Main.ReportNo;
                worksheet.Cells[4, 4] = model.Main.ReportNo;
                worksheet.Cells[4, 7] = model.Main.OrderID;

                worksheet.Cells[5, 4] = model.Main.FactoryID;
                worksheet.Cells[5, 7] = model.Main.SubmitDateText;

                worksheet.Cells[6, 4] = model.Main.StyleID;
                worksheet.Cells[6, 7] = model.Main.ReportDateText;

                worksheet.Cells[7, 4] = model.Main.Article;
                worksheet.Cells[8, 4] = model.Main.SeasonID;
                worksheet.Cells[9, 4] = model.Main.FabricRefNo;
                worksheet.Cells[10, 4] = model.Main.FabricColor;
                worksheet.Cells[11, 5] = model.Main.FabricDescription;


                // Technician 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["Technician"].ToString();

                    // 姓名
                    worksheet.Cells[69, 6] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[69, 7];
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

                // TestBeforePicture、TestAfterPicture 圖片
                if (model.Main.TestBeforePicture != null && model.Main.TestBeforePicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.get_Range("B21:F40");
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestBeforePicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, cell.Width - 10, cell.Height - 10);

                    cell = worksheet.get_Range("B42:F61");
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, cell.Width - 10, cell.Height - 10);
                }
                if (model.Main.TestBeforeWashPicture != null && model.Main.TestBeforeWashPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.get_Range("G21:I40");
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestBeforeWashPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, cell.Width - 10, cell.Height - 10);
                }
                if (model.Main.TestAfterPicture != null && model.Main.TestAfterPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.get_Range("G42:I61");
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestAfterPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, cell.Width - 10, cell.Height - 10);
                }

                // 表身處理
                if (model.DetailList.Any() && model.DetailList.Count >= 1)
                {

                    WaterAbsorbency_Detail detailData = model.DetailList.Where(x => x.EvaluationItem == "Before Wash").FirstOrDefault();
                    worksheet.Cells[15, 4] = detailData.NoOfDrops;
                    if (model.Main.FabricType == "KNIT")
                    {
                        worksheet.Cells[15, 5] = detailData.Values;
                    }
                    else if (model.Main.FabricType == "WOVEN")
                    {
                        worksheet.Cells[15, 7] = detailData.Values;
                    }
                    worksheet.Cells[15, 8] = detailData.Result;
                    worksheet.Cells[15, 9] = model.Main.Result;

                    detailData = model.DetailList.Where(x => x.EvaluationItem == "After 5 Cycle Wash").FirstOrDefault();
                    worksheet.Cells[17, 4] = detailData.NoOfDrops;
                    if (model.Main.FabricType == "KNIT")
                    {
                        worksheet.Cells[17, 5] = detailData.Values;
                    }
                    else if (model.Main.FabricType == "WOVEN")
                    {
                        worksheet.Cells[17, 7] = detailData.Values;
                    }
                    worksheet.Cells[17, 8] = detailData.Result;
                }

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }

                string fileName = $"{tmpName}.xlsx";
                string fullExcelFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                string filePdfName = $"{tmpName}.pdf";
                string fullPdfFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filePdfName);


                Microsoft.Office.Interop.Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(fullExcelFileName);

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);

                // 轉PDF再繼續進行以下
                if (isPDF)
                {
                    if (ConvertToPDF.ExcelToPDF(fullExcelFileName, fullPdfFileName))
                    {
                        result.TempFileName = filePdfName;
                        result.Result = true;
                    }
                    else
                    {
                        result.ErrorMessage = "Convert To PDF Fail";
                        result.Result = false;
                    }
                }
                else
                {
                    result.TempFileName = fileName;
                    result.Result = true;
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.Result = false;
            }
            finally
            {
                Marshal.ReleaseComObject(excel);
            }
            return result;
        }

        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new WaterAbsorbencyProvider(Common.ManufacturingExecutionDataAccessLayer);

            WaterAbsorbency_ViewModel model = this.GetData(new WaterAbsorbency_Request() { ReportNo = ReportNo });
            string name = $"Water Absorbency Test_{model.Main.OrderID}_" +
                $"{model.Main.StyleID}_" +
                $"{model.Main.FabricRefNo}_" +
                $"{model.Main.FabricColor}_" +
                $"{model.Main.Result}_" +
                $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            WaterAbsorbency_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Water Absorbency Test/{model.Main.OrderID}/" +
                $"{model.Main.StyleID}/" +
                $"{model.Main.FabricRefNo}/" +
                $"{model.Main.FabricColor}/" +
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
