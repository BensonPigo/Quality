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
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class WickingHeightTestService
    {
        private WickingHeightTestProvider _Provider;
        private MailToolsService _MailService;

        public WickingHeightTest_ViewModel GetDefaultModel(bool iNew = false)
        {
            WickingHeightTest_ViewModel model = new WickingHeightTest_ViewModel()
            {
                Request = new WickingHeightTest_Request(),
                Main = new WickingHeightTest_Main(),
                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<WickingHeightTest_Detail>() 
                { 
                    new WickingHeightTest_Detail() { EvaluationType = "Before Wash" },
                    new WickingHeightTest_Detail() { EvaluationType = "After Wash" },
                },
                DetaiItemlList = new List<WickingHeightTest_Detail_Item>()
                {
                    new WickingHeightTest_Detail_Item() { EvaluationType = "Before Wash", EvaluationItem = "Specimen 1", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "Before Wash", EvaluationItem = "Specimen 2", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "Before Wash", EvaluationItem = "Specimen 3", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "After Wash", EvaluationItem = "Specimen 1", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "After Wash", EvaluationItem = "Specimen 2", WarpTime = 30, WeftTime = 30 },
                    new WickingHeightTest_Detail_Item() { EvaluationType = "After Wash", EvaluationItem = "Specimen 3", WarpTime = 30, WeftTime = 30 }
                },
            };

            try
            {
                _Provider = new WickingHeightTestProvider(Common.ProductionDataAccessLayer);

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
        public WickingHeightTest_ViewModel GetData(WickingHeightTest_Request Req)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new WickingHeightTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<WickingHeightTest_Main> tmpList = new List<WickingHeightTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new WickingHeightTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new WickingHeightTest_Request()
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
                    _Provider = new WickingHeightTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new WickingHeightTest_Request() { OrderID = model.Main.OrderID });

                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    _Provider = new WickingHeightTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new WickingHeightTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    model.DetaiItemlList = _Provider.GetDetaiItemlList(new WickingHeightTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

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

        public WickingHeightTest_ViewModel GetOrderInfo(string OrderID)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new WickingHeightTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new WickingHeightTest_Request() { OrderID = OrderID });


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

        public WickingHeightTest_ViewModel NewSave(WickingHeightTest_ViewModel Req, string MDivision, string UserID)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WickingHeightTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.WarpResult == "Fail" || o.WeftResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<WickingHeightTest_Detail>();
                }


                // 新增，並取得ReportNo
                _Provider.Insert_WickingHeightTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // WickingHeightTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<WickingHeightTest_Detail>();
                }

                // WickingHeightTest_Detail 新增 or 修改
                if (Req.DetaiItemlList == null)
                {
                    Req.DetaiItemlList = new List<WickingHeightTest_Detail_Item>();
                }

                _Provider.Processe_WickingHeightTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WickingHeightTest_Request()
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

        public WickingHeightTest_ViewModel EditSave(WickingHeightTest_ViewModel Req, string UserID)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WickingHeightTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.WarpResult == "Fail" || o.WeftResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<WickingHeightTest_Detail>();
                }


                // 更新表頭，並取得ReportNo
                _Provider.Update_WickingHeightTest(Req, UserID);

                // 更新表身
                _Provider.Processe_WickingHeightTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WickingHeightTest_Request()
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

        public WickingHeightTest_ViewModel Delete(WickingHeightTest_ViewModel Req)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WickingHeightTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_WickingHeightTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new WickingHeightTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new WickingHeightTest_Request()
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

        public WickingHeightTest_ViewModel EncodeAmend(WickingHeightTest_ViewModel Req, string UserID)
        {
            WickingHeightTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new WickingHeightTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_WickingHeightTest(Req.Main, UserID);

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

        public WickingHeightTest_ViewModel GetReport(string ReportNo, bool isPDF)
        {
            WickingHeightTest_ViewModel result = new WickingHeightTest_ViewModel();

            string basefileName = "WickingHeightTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

            Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);
            try
            {
                _Provider = new WickingHeightTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得報表資料

                WickingHeightTest_ViewModel model = this.GetData(new WickingHeightTest_Request() { ReportNo = ReportNo });

                DataTable ReportTechnician = _Provider.GetReportTechnician(new WickingHeightTest_Request() { ReportNo = ReportNo });

                excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表

                // 取得工作表上所有圖形物件
                Microsoft.Office.Interop.Excel.Shapes shapes = worksheet.Shapes;

                string reportNo = model.Main.ReportNo;
                worksheet.Cells[2, 3] = model.Main.ReportNo;
                worksheet.Cells[2, 6] = model.Main.ReportDate;

                worksheet.Cells[3, 3] = model.Main.SubmitDate;
                worksheet.Cells[3, 6] = model.Main.BrandID;

                worksheet.Cells[4, 3] = model.Main.SeasonID;
                worksheet.Cells[4, 6] = model.Main.OrderID;

                worksheet.Cells[5, 3] = model.Main.StyleID;
                worksheet.Cells[5, 6] = "Bulk";

                worksheet.Cells[6, 3] = model.Main.Article;
                worksheet.Cells[6, 6] = model.Main.FabricType;

                worksheet.Cells[7, 2] = model.Main.Result;
                worksheet.Cells[7, 5] = model.Main.FabricDescription;

                // Technician 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["Technician"].ToString();

                    // 姓名
                    worksheet.Cells[23, 3] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[24, 3];
                    if (ReportTechnician.Rows[0]["TechnicianSignture"] != DBNull.Value)
                    {

                        byte[] TestBeforePicture = (byte[])ReportTechnician.Rows[0]["TechnicianSignture"]; // 圖片的 byte[]

                        MemoryStream ms = new MemoryStream(TestBeforePicture);
                        System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                        string imageName = $"{Guid.NewGuid()}.jpg";
                        string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                        img.Save(imgPath);
                        worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, cell.Width, cell.Height);

                    }
                }


                // TestWarpPicture 圖片
                if (model.Main.TestWarpPicture != null && model.Main.TestWarpPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[21, 1];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestWarpPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

                // TestAfterPicture 圖片
                if (model.Main.TestWeftPicture != null && model.Main.TestWeftPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[21, 4];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestWeftPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }


                // 表身處理
                if (model.DetailList.Any() && model.DetailList.Count >= 1)
                {
                    long detailUkey = model.DetailList.Where(x => x.EvaluationType == "Before Wash").Select(x => x.Ukey).FirstOrDefault();
                    int i = 0;
                    foreach(var item in model.DetaiItemlList.Where(x => x.WickingHeightTestDetailUkey == detailUkey).OrderBy(x => x.Ukey))
                    {
                        worksheet.Cells[14, 2 + i] = string.Format("{0}mm/{1}min", item.WarpValues.ToString(), item.WarpTime.ToString());
                        worksheet.Cells[14, 5 + i] = string.Format("{0}mm/{1}min", item.WeftValues.ToString(), item.WeftTime.ToString());
                        i++;
                    }

                    detailUkey = model.DetailList.Where(x => x.EvaluationType == "After Wash").Select(x => x.Ukey).FirstOrDefault();
                    i = 0;
                    foreach (var item in model.DetaiItemlList.Where(x => x.WickingHeightTestDetailUkey == detailUkey).OrderBy(x => x.Ukey))
                    {
                        worksheet.Cells[18, 2 + i] = string.Format("{0}mm/{1}min", item.WarpValues.ToString(), item.WarpTime.ToString());
                        worksheet.Cells[18, 5 + i] = string.Format("{0}mm/{1}min", item.WeftValues.ToString(), item.WeftTime.ToString());
                        i++;
                    }
                }

                string fileName = $"WickingHeightTest_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string fullExcelFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                string filePdfName = $"WickingHeightTest_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
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
        public SendMail_Result SendMail(string ReportNo, string TO, string CC)
        {
            _Provider = new WickingHeightTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            WickingHeightTest_ViewModel model = this.GetData(new WickingHeightTest_Request() { ReportNo = ReportNo });

            WickingHeightTest_ViewModel report = this.GetReport(ReportNo, false);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Wicking Height Test/{model.Main.OrderID}/" +
                    $"{model.Main.StyleID}/" +
                    $"{model.Main.Article}/" +
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
