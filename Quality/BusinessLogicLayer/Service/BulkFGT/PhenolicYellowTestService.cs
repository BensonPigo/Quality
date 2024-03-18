using ADOHelper.Utility;
using BusinessLogicLayer.Helper;
using DatabaseObject.ProductionDB;
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
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using static Sci.MyUtility;
using static System.Net.Mime.MediaTypeNames;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class PhenolicYellowTestService
    {
        private PhenolicYellowTestProvider _Provider;
        private MailToolsService _MailService;
        public PhenolicYellowTest_ViewModel GetDefaultModel(bool iNew = false)
        {
            PhenolicYellowTest_ViewModel model = new PhenolicYellowTest_ViewModel()
            {
                Request = new PhenolicYellowTest_Request(),
                Main = new PhenolicYellowTest_Main(),
                
                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                Scale_Source = new List<System.Web.Mvc.SelectListItem>(),
                Temperature_Source = new List<System.Web.Mvc.SelectListItem>()
                {
                    new System.Web.Mvc.SelectListItem(){Text="0",Value="0"},
                    new System.Web.Mvc.SelectListItem(){Text="30",Value="30"},
                    new System.Web.Mvc.SelectListItem(){Text="40",Value="40"},
                    new System.Web.Mvc.SelectListItem(){Text="50",Value="50"},
                    new System.Web.Mvc.SelectListItem(){Text="60",Value="60"},
                },
                DetailList = new List<PhenolicYellowTest_Detail>(),
            };

            try
            {
                _Provider = new PhenolicYellowTestProvider(Common.ProductionDataAccessLayer);
                model.Scale_Source = _Provider.GetScales();

                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }
            return model;
        }

        /// <summary>
        /// 取得查詢結果
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public PhenolicYellowTest_ViewModel GetData(PhenolicYellowTest_Request Req)
        {
            PhenolicYellowTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new PhenolicYellowTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<PhenolicYellowTest_Main> tmpList = new List<PhenolicYellowTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new PhenolicYellowTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new PhenolicYellowTest_Request()
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
                    if (!string.IsNullOrEmpty(Req.ReportNo) && tmpList.Where(o => o.ReportNo == Req.ReportNo).Any())
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
                    _Provider = new PhenolicYellowTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new PhenolicYellowTest_Request() { OrderID = model.Main.OrderID });

                    foreach (var oriData in tmpOrders)
                    {
                        SelectListItem Article = new SelectListItem()
                        {
                            Text = oriData.Article,
                            Value = oriData.Article,
                        };
                        model.Article_Source.Add(Article);
                    }

                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    _Provider = new PhenolicYellowTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new PhenolicYellowTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    string Subject = $"Phenolic Yellowing Test/{model.Main.OrderID}/" +
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
        public PhenolicYellowTest_ViewModel GetOrderInfo(string OrderID)
        {
            PhenolicYellowTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new PhenolicYellowTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new PhenolicYellowTest_Request() { OrderID = OrderID });


                // 確認SP#是否存在
                if (tmpOrders.Any())
                {
                    // 取得表頭SP#相關欄位
                    model.Main.OrderID = tmpOrders.FirstOrDefault().ID;
                    model.Main.FactoryID = tmpOrders.FirstOrDefault().FactoryID;
                    model.Main.BrandID = tmpOrders.FirstOrDefault().BrandID;
                    model.Main.SeasonID = tmpOrders.FirstOrDefault().SeasonID;
                    model.Main.StyleID = tmpOrders.FirstOrDefault().StyleID;

                    // 取得Article 下拉選單
                    foreach (var oriData in tmpOrders)
                    {
                        System.Web.Mvc.SelectListItem Article = new System.Web.Mvc.SelectListItem()
                        {
                            Text = oriData.Article,
                            Value = oriData.Article,
                        };
                        model.Article_Source.Add(Article);
                    }

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
        public PhenolicYellowTest_ViewModel NewSave(PhenolicYellowTest_ViewModel Req, string MDivision, string UserID)
        {
            PhenolicYellowTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new PhenolicYellowTestProvider(_ISQLDataTransaction);

                if (string.IsNullOrEmpty(Req.Main.OrderID) || string.IsNullOrEmpty(Req.Main.BrandID) || string.IsNullOrEmpty(Req.Main.SeasonID)|| string.IsNullOrEmpty(Req.Main.StyleID) || string.IsNullOrEmpty(Req.Main.Article))
                {
                    throw new Exception("SP#、Brand、Season、Style and Article can't be empty.");
                }

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<PhenolicYellowTest_Detail>();
                }

                // 新增，並取得ReportNo
                _Provider.Insert_PhenolicYellowTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // PhenolicYellowTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<PhenolicYellowTest_Detail>();
                }

                _Provider.Processe_PhenolicYellowTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new PhenolicYellowTest_Request()
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
        public PhenolicYellowTest_ViewModel EditSave(PhenolicYellowTest_ViewModel Req, string UserID)
        {
            PhenolicYellowTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new PhenolicYellowTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.Result == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<PhenolicYellowTest_Detail>();
                }


                // 更新表頭，並取得ReportNo
                _Provider.Update_PhenolicYellowTest(Req, UserID);

                // 更新表身
                _Provider.Processe_PhenolicYellowTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new PhenolicYellowTest_Request()
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
        public PhenolicYellowTest_ViewModel Delete(PhenolicYellowTest_ViewModel Req)
        {
            PhenolicYellowTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new PhenolicYellowTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_PhenolicYellowTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new PhenolicYellowTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new PhenolicYellowTest_Request()
                {
                    ReportNo = model.Main.ReportNo,
                    BrandID =model.Main.BrandID,
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
        public PhenolicYellowTest_ViewModel EncodeAmend(PhenolicYellowTest_ViewModel Req, string UserID)
        {
            PhenolicYellowTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new PhenolicYellowTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_PhenolicYellowTest(Req.Main, UserID);

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

        public PhenolicYellowTest_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            PhenolicYellowTest_ViewModel result = new PhenolicYellowTest_ViewModel();

            string basefileName = "PhenolicYellowTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";
            string tmpName = string.Empty;

            Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);

            try
            {
                _Provider = new PhenolicYellowTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得報表資料

                PhenolicYellowTest_ViewModel model = this.GetData(new PhenolicYellowTest_Request() { ReportNo = ReportNo });

                DataTable ReportTechnician = _Provider.GetReportTechnician(new PhenolicYellowTest_Request() { ReportNo = ReportNo });

                tmpName = $"Phenolic Yellowing Test_{model.Main.OrderID}_" +
                    $"{model.Main.StyleID}_" +
                    $"{model.Main.FabricRefNo}_" +
                    $"{model.Main.FabricColor}_" +
                    $"{model.Main.Result}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表

                string reportNo = model.Main.ReportNo;
                worksheet.Cells[3, 2] = model.Main.ReportNo;
                worksheet.Cells[3, 5] = model.Main.SubmitDateText;

                worksheet.Cells[4, 2] = model.Main.OrderID;
                worksheet.Cells[4, 5] = model.Main.ReportDateText;

                worksheet.Cells[5, 2] = model.Main.BrandID;
                worksheet.Cells[5, 5] = model.Main.SeasonID;

                worksheet.Cells[6, 2] = model.Main.StyleID;
                worksheet.Cells[6, 5] = model.Main.Article;

                worksheet.Cells[7, 2] = model.Main.Seq;
                worksheet.Cells[7, 5] = model.Main.FabricRefNo;

                worksheet.Cells[8, 2] = model.Main.FabricColor;
                worksheet.Cells[8, 5] = model.Main.Result;

                worksheet.Cells[11, 2] = model.Main.Temperature;
                worksheet.Cells[11, 5] = model.Main.Time;

                // Technician 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["Technician"].ToString();

                    // 姓名
                    worksheet.Cells[24, 6] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[23, 6];
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
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[21, 1];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestBeforePicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }
                if (model.Main.TestAfterPicture != null && model.Main.TestAfterPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[21, 4];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestAfterPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

                // 表身處理 單筆和多筆分開
                if (model.DetailList.Any() && model.DetailList.Count > 1)
                {
                    // 先處理Remark
                    List<string> allRemark = new List<string>();
                    allRemark = model.DetailList.Where(o => !string.IsNullOrEmpty(o.Remark)).Select(o => o.Remark).ToList();

                    // 全部擠在一起，但是要分行
                    worksheet.Cells[17, 2] = string.Join(Environment.NewLine, allRemark);

                    // 複製欄位
                    int copyCount = model.DetailList.Count - 1;
                    for (int i = 0; i < copyCount; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range paste1 = worksheet.get_Range($"A{i + 15}", Type.Missing);
                        Microsoft.Office.Interop.Excel.Range copyRow = worksheet.get_Range("A15").EntireRow;
                        paste1.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }

                    int rowIdx = 0;
                    foreach (var detailData in model.DetailList)
                    {
                        worksheet.Cells[15 + rowIdx, 1] = detailData.EvaluationItem;
                        worksheet.Cells[15 + rowIdx, 2] = detailData.Dyelot;
                        worksheet.Cells[15 + rowIdx, 3] = detailData.Roll;
                        worksheet.Cells[15 + rowIdx, 5] = detailData.Scale;
                        worksheet.Cells[15 + rowIdx, 6] = detailData.Result;

                        allRemark.Add(detailData.Remark);
                        rowIdx++;
                    }
                }
                else
                {
                    foreach (var detailData in model.DetailList)
                    {
                        worksheet.Cells[15, 1] = detailData.EvaluationItem;
                        worksheet.Cells[15, 2] = detailData.Dyelot;
                        worksheet.Cells[15, 3] = detailData.Roll;
                        worksheet.Cells[15, 5] = detailData.Scale;
                        worksheet.Cells[15, 6] = detailData.Result;
                        worksheet.Cells[17, 2] = detailData.Remark;
                    }
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
            _Provider = new PhenolicYellowTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            PhenolicYellowTest_ViewModel model = this.GetData(new PhenolicYellowTest_Request() { ReportNo = ReportNo });

            string name = $"Phenolic Yellowing Test_{model.Main.OrderID}_" +
                    $"{model.Main.StyleID}_" +
                    $"{model.Main.FabricRefNo}_" +
                    $"{model.Main.FabricColor}_" +
                    $"{model.Main.Result}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            PhenolicYellowTest_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Phenolic Yellowing Test/{model.Main.OrderID}/" +
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
