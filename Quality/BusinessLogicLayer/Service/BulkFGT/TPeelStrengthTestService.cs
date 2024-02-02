using ADOHelper.Utility;
using DatabaseObject.ViewModel.BulkFGT;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using iTextSharp.text;
using iTextSharp.text.pdf;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class TPeelStrengthTestService
    {
        private string UploadFileRootPath = ConfigurationManager.AppSettings["UploadFileRootPath"];
        private string UploadFilePath = $@"{ConfigurationManager.AppSettings["UploadFileRootPath"]}BulkFGT\TPeelStrengthTest\";
        private TPeelStrengthTestProvider _Provider;
        private MailToolsService _MailService;

        public TPeelStrengthTest_ViewModel GetDefaultModel(bool iNew = false)
        {
            TPeelStrengthTest_ViewModel model = new TPeelStrengthTest_ViewModel()
            {
                Request = new TPeelStrengthTest_Request(),
                Main = new TPeelStrengthTest_Main()
                {
                    // ISP20230288 寫死4
                    MachineReport = "",

                },
                Article_Source = new List<System.Web.Mvc.SelectListItem>(),
                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<TPeelStrengthTest_Detail>(),

                MachineNo_Source = new List<System.Web.Mvc.SelectListItem>()
                {
                    // ISP20230288 寫死37
                    new System.Web.Mvc.SelectListItem(){Text="Fabric",Value="Fabric"},
                    new System.Web.Mvc.SelectListItem(){Text="Accessory",Value="Accessory"},
                    new System.Web.Mvc.SelectListItem(){Text="Printing",Value="Printing"},
                },
            };


            if (!System.IO.Directory.Exists($@"{UploadFileRootPath}\BulkFGT\"))
            {
                System.IO.Directory.CreateDirectory($@"{UploadFileRootPath}\BulkFGT\");
            }

            if (!System.IO.Directory.Exists($@"{UploadFileRootPath}\BulkFGT\TPeelStrengthTest"))
            {
                System.IO.Directory.CreateDirectory($@"{UploadFileRootPath}\BulkFGT\TPeelStrengthTest");
            }

            return model;
        }

        /// <summary>
        /// 取得查詢結果
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public TPeelStrengthTest_ViewModel GetData(TPeelStrengthTest_Request Req)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new TPeelStrengthTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<TPeelStrengthTest_Main> tmpList = new List<TPeelStrengthTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new TPeelStrengthTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new TPeelStrengthTest_Request()
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
                    _Provider = new TPeelStrengthTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new TPeelStrengthTest_Request() { OrderID = model.Main.OrderID });

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

                    _Provider = new TPeelStrengthTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new TPeelStrengthTest_Request()
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

        public TPeelStrengthTest_ViewModel GetOrderInfo(string OrderID)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new TPeelStrengthTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new TPeelStrengthTest_Request() { OrderID = OrderID });


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

        public TPeelStrengthTest_ViewModel NewSave(TPeelStrengthTest_ViewModel Req, string MDivision, string UserID)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new TPeelStrengthTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.AllResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<TPeelStrengthTest_Detail>();
                }

                if (Req.Main.MachineReportFile != null)
                {
                    this.CreateMachineReportFile(Req);
                }

                // 新增，並取得ReportNo
                _Provider.Insert_TPeelStrengthTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // TPeelStrengthTest_Detail 新增 or 修改
                if (Req.DetailList == null)
                {
                    Req.DetailList = new List<TPeelStrengthTest_Detail>();
                }

                _Provider.Processe_TPeelStrengthTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new TPeelStrengthTest_Request()
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

        public TPeelStrengthTest_ViewModel EditSave(TPeelStrengthTest_ViewModel Req, string UserID)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new TPeelStrengthTestProvider(_ISQLDataTransaction);

                // 判斷表頭Result，表身有任一Fail則Fail，否則Pass
                if (Req.DetailList != null && Req.DetailList.Any())
                {
                    bool HasFail = Req.DetailList.Where(o => o.AllResult == "Fail").Any();
                    Req.Main.Result = HasFail ? "Fail" : "Pass";
                }
                else
                {
                    Req.Main.Result = "Pass";
                    Req.DetailList = new List<TPeelStrengthTest_Detail>();
                }

                if (Req.Main.MachineReportFile != null)
                {
                    this.CreateMachineReportFile(Req);
                }

                // 更新表頭
                _Provider.Update_TPeelStrengthTest(Req, UserID);

                // 更新表身
                _Provider.Processe_TPeelStrengthTest_Detail(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new TPeelStrengthTest_Request()
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
        public TPeelStrengthTest_ViewModel Delete(TPeelStrengthTest_ViewModel Req)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new TPeelStrengthTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_TPeelStrengthTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new TPeelStrengthTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new TPeelStrengthTest_Request()
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
        public TPeelStrengthTest_ViewModel EncodeAmend(TPeelStrengthTest_ViewModel Req, string UserID)
        {
            TPeelStrengthTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new TPeelStrengthTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_TPeelStrengthTest(Req.Main, UserID);

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

        public TPeelStrengthTest_ViewModel GetReport(string ReportNo, bool isPDF)
        {
            TPeelStrengthTest_ViewModel result = new TPeelStrengthTest_ViewModel();

            string basefileName = "TPeelStrengthTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

            Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);

            try
            {
                _Provider = new TPeelStrengthTestProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得報表資料

                TPeelStrengthTest_ViewModel model = this.GetData(new TPeelStrengthTest_Request() { ReportNo = ReportNo });

                DataTable ReportTechnician = _Provider.GetReportTechnician(new TPeelStrengthTest_Request() { ReportNo = ReportNo });

                excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.Sheets[1]; // 取得工作表


                string reportNo = model.Main.ReportNo;
                string machineReport = string.IsNullOrEmpty(model.Main.MachineReport) ? string.Empty : model.Main.MachineReport;

                worksheet.Cells[3, 2] = model.Main.ReportNo;

                worksheet.Cells[4, 2] = model.Main.SubmitDateText;
                worksheet.Cells[4, 5] = model.Main.ReportDateText;

                worksheet.Cells[5, 2] = model.Main.OrderID;
                worksheet.Cells[5, 5] = model.Main.BrandID;

                worksheet.Cells[6, 2] = model.Main.StyleID;
                worksheet.Cells[6, 5] = model.Main.SeasonID;

                worksheet.Cells[7, 2] = model.Main.Article;
                worksheet.Cells[7, 5] = model.Main.MachineNo;

                worksheet.Cells[8, 2] = model.Main.FabricRefNo;
                worksheet.Cells[8, 5] = model.Main.FabricColor;


                // Technician 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["Technician"].ToString();

                    // 姓名
                    worksheet.Cells[20, 6] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[19, 6];
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

                // TestAfterPicture 圖片
                if (model.Main.TestAfterPicture != null && model.Main.TestAfterPicture.Length > 1)
                {
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[17, 2];
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(model.Main.TestAfterPicture, reportNo, ToolKit.PublicClass.SingLocation.MiddleItalic, test: false);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left + 5, cell.Top + 5, 200, 300);
                }

                // 表身處理
                if (model.DetailList.Any() && model.DetailList.Count > 1)
                {
                    //// 先處理Remark
                    string allRemark = string.Join(Environment.NewLine, model.DetailList.Select(o => o.Remark));

                    //// 全部擠在一起，但是要分行
                    worksheet.Cells[14, 2] = allRemark;

                    // 複製欄位
                    int copyCount = model.DetailList.Count - 1;
                    for (int i = 0; i < copyCount; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range paste1 = worksheet.get_Range($"A12", Type.Missing);
                        Microsoft.Office.Interop.Excel.Range copyRow = worksheet.get_Range("12:12");
                        paste1.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }

                    int rowIdx = 0;
                    foreach (var detailData in model.DetailList)
                    {
                        // Wash Before or Wash After
                        worksheet.Cells[12 + rowIdx, 1] = $@"Wash {detailData.EvaluationItem}";

                        if (detailData.EvaluationItem.ToLower() == "before")
                        {
                            worksheet.Cells[12 + rowIdx, 3] = detailData.WarpValue;
                            worksheet.Cells[12 + rowIdx, 4] = detailData.WarpResult;
                            worksheet.Cells[12 + rowIdx, 5] = detailData.WeftValue;
                            worksheet.Cells[12 + rowIdx, 6] = detailData.WeftResult;
                        }
                        if (detailData.EvaluationItem.ToLower() == "after")
                        {
                            worksheet.Cells[12 + rowIdx, 3] = detailData.WarpValue;
                            worksheet.Cells[12 + rowIdx, 4] = detailData.WarpResult;
                            worksheet.Cells[12 + rowIdx, 5] = detailData.WeftValue;
                            worksheet.Cells[12 + rowIdx, 6] = detailData.WeftResult;
                        }

                        //worksheet.Cells[15 + rowIdx, 2] = detailData.Remark;
                        rowIdx += 1;
                    }
                }

                string fileName = $"TPeelStrengthTest_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string fullExcelFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                string filePdfName = $"TPeelStrengthTest_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
                string fullPdfFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filePdfName);


                Microsoft.Office.Interop.Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(fullExcelFileName);

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);

                if (ConvertToPDF.ExcelToPDF(fullExcelFileName, fullPdfFileName))
                {
                    string outputPdfName = $"TPeelStrengthTest_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
                    string outputPdfFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", outputPdfName);
                    string pdf1Path = fullPdfFileName;


                    if (!string.IsNullOrEmpty(machineReport))
                    {
                        string machineReportFile = this.UploadFilePath + machineReport;
                        string pdf2Path = machineReportFile;

                        // 建立一個新的 PDF 文件
                        Document mergedDocument = new Document();

                        // 建立一個 PDFWriter 來寫入合併後的 PDF
                        PdfWriter writer = PdfWriter.GetInstance(mergedDocument, new FileStream(outputPdfFileName, FileMode.Create));

                        // 開啟合併後的 PDF 文件
                        mergedDocument.Open();

                        // 合併第一個 PDF
                        PdfReader pdf1 = new PdfReader(pdf1Path);
                        CopyPages(pdf1, mergedDocument, writer);

                        // 合併第二個 PDF
                        PdfReader pdf2 = new PdfReader(pdf2Path);
                        CopyPages(pdf2, mergedDocument, writer);

                        // 關閉合併後的 PDF 文件
                        mergedDocument.Close();

                        // 釋放資源
                        pdf1.Close();
                        pdf2.Close();

                        result.TempFileName = outputPdfName;
                    }
                    else
                    {
                        result.TempFileName = filePdfName;
                    }
                    result.Result = true;


                }
                else
                {
                    result.ErrorMessage = "Convert To PDF Fail";
                    result.Result = false;
                }


                //// 轉PDF再繼續進行以下
                //if (isPDF)
                //{
                //    if (ConvertToPDF.ExcelToPDF(fullExcelFileName, fullPdfFileName))
                //    {
                //        result.TempFileName = filePdfName;
                //        result.Result = true;
                //    }
                //    else
                //    {
                //        result.ErrorMessage = "Convert To PDF Fail";
                //        result.Result = false;
                //    }
                //}
                //else
                //{
                //    result.TempFileName = fileName;
                //    result.Result = true;
                //}
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

        // 定義一個方法來將頁面從一個 PDF 複製到另一個 PDF
        public void CopyPages(PdfReader sourcePdf, Document destinationPdf, PdfWriter writer)
        {
            PdfContentByte contentByte = writer.DirectContent;
            for (int pageIndex = 1; pageIndex <= sourcePdf.NumberOfPages; pageIndex++)
            {
                PdfImportedPage importedPage = writer.GetImportedPage(sourcePdf, pageIndex);
                destinationPdf.NewPage();
                contentByte.AddTemplate(importedPage, 0, 0);
            }
        }

        public void CreateMachineReportFile(TPeelStrengthTest_ViewModel Req)
        {
            Req.Main.MachineReport = Req.Main.MachineReportFile.FileName;
            //string FileName= Req.Main.MachineReportFile.FileName;
            string FileNameWithoutExtension = Path.GetFileNameWithoutExtension(Req.Main.MachineReportFile.FileName);
            string FileExtension = Path.GetExtension(Req.Main.MachineReportFile.FileName);
            string FullFileName = $@"{this.UploadFilePath}\{Req.Main.MachineReport}";

            int count = 1;
            string FileName = FileNameWithoutExtension;
            while (File.Exists(FullFileName))
            {
                FileName = $"{FileNameWithoutExtension} ({count})";
                FullFileName = Path.Combine(this.UploadFilePath, $"{FileName}{FileExtension}");
                count++;
            }

            using (var fileStream = File.Create(FullFileName))
            {
                Req.Main.MachineReportFile.InputStream.Seek(0, SeekOrigin.Begin);
                Req.Main.MachineReportFile.InputStream.CopyTo(fileStream);
            }

            Req.Main.MachineReport = FileName + FileExtension;
        }
        public SendMail_Result SendMail(string ReportNo, string TO, string CC)
        {
            _Provider = new TPeelStrengthTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            TPeelStrengthTest_ViewModel model = this.GetData(new TPeelStrengthTest_Request() { ReportNo = ReportNo });

            TPeelStrengthTest_ViewModel report = this.GetReport(ReportNo, false);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"T-Peel Strength  Test/{model.Main.OrderID}/" +
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
