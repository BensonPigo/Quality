using ADOHelper.Utility;
using BusinessLogicLayer.Helper;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using Ict;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Org.BouncyCastle.Ocsp;
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
using System.Web.Mvc;
using static Sci.MyUtility;
using static System.Net.Mime.MediaTypeNames;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class BrandGarmentTestService
    {
        private string UploadFileRootPath = ConfigurationManager.AppSettings["UploadFileRootPath"];
        private string UploadFilePath = $@"{ConfigurationManager.AppSettings["UploadFileRootPath"]}BulkFGT\BrandGarmentTest\";

        private BrandGarmentTestProvider _Provider;
        public BrandGarmentTest_ViewModel GetDefaultModel(bool iNew = false)
        {
            BrandGarmentTest_ViewModel model = new BrandGarmentTest_ViewModel()
            {
                Request = new BrandGarmentTest_Request(),
                Main = new BrandGarmentTest()
                {
                    Result = "Pass"
                },
                MainList = new List<BrandGarmentTest>(),
                BrandGarmentTestDoxList = new List<BrandGarmentTestDox>(),
                Article_Source = new List<SelectListItem>(),
            };

            return model;
        }

        /// <summary>
        /// 取得查詢結果
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public BrandGarmentTest_ViewModel GetMainList(BrandGarmentTest_Request Req)
        {
            BrandGarmentTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new BrandGarmentTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<BrandGarmentTest> tmpList = new List<BrandGarmentTest>();

                // 取得列表資料
                tmpList = _Provider.GetMainList(Req);


                if (tmpList.Any())
                {
                    model.MainList = tmpList;
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
        public BrandGarmentTest_ViewModel GetMain(BrandGarmentTest_Request Req)
        {
            BrandGarmentTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new BrandGarmentTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<BrandGarmentTest> tmpList = new List<BrandGarmentTest>();

                // 取得列表資料
                tmpList = _Provider.GetMainList(Req);


                if (tmpList.Any())
                {
                    model.Main = tmpList.FirstOrDefault();

                    List<BrandGarmentTestDox> doxList = _Provider.GetBrandGarmentTestDoxList(new BrandGarmentTest_Request() { ReportNo = model.Main.ReportNo });

                    model.BrandGarmentTestDoxList = doxList;

                    // 取得Article 下拉選單
                    _Provider = new BrandGarmentTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new BrandGarmentTest_Request() { OrderID = model.Main.OrderID });

                    foreach (var oriData in tmpOrders)
                    {
                        SelectListItem Article = new SelectListItem()
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
        public BrandGarmentTest_ViewModel GetOrderInfo(string OrderID)
        {
            BrandGarmentTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new BrandGarmentTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new BrandGarmentTest_Request() { OrderID = OrderID });


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
        public BrandGarmentTest_ViewModel NewSave(BrandGarmentTest_ViewModel Req, string MDivision, string UserID)
        {
            BrandGarmentTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new BrandGarmentTestProvider(_ISQLDataTransaction);

                if (string.IsNullOrEmpty(Req.Main.OrderID) || string.IsNullOrEmpty(Req.Main.BrandID) || string.IsNullOrEmpty(Req.Main.SeasonID) || string.IsNullOrEmpty(Req.Main.StyleID) || string.IsNullOrEmpty(Req.Main.Article))
                {
                    throw new Exception("SP#、Brand、Season、Style and Article can't be empty.");
                }


                // 新增，並取得ReportNo
                _Provider.Insert_BrandGarmentTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // BrandGarmentTestDoxList 檔案上傳
                if (Req.BrandGarmentTestDoxList == null || !Req.BrandGarmentTestDoxList.Any())
                {
                    Req.BrandGarmentTestDoxList = new List<BrandGarmentTestDox>();
                }
                else
                {
                    foreach (var dox in Req.BrandGarmentTestDoxList.Where(o => o.IsOldFile == false))
                    {
                        dox.FilePath = this.UploadFilePath;
                        this.CreateMachineReportFile(dox);
                    }
                }


                _Provider.Processe_BrandGarmentTestDox(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model.Main = this.GetMain(new BrandGarmentTest_Request()
                {
                    ReportNo = Req.Main.ReportNo,
                }).Main;

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
        public BrandGarmentTest_ViewModel EditSave(BrandGarmentTest_ViewModel Req, string UserID)
        {
            BrandGarmentTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new BrandGarmentTestProvider(_ISQLDataTransaction);


                // 更新表頭，並取得ReportNo
                _Provider.Update_BrandGarmentTest(Req, UserID);


                // BrandGarmentTestDoxList 檔案上傳
                if (Req.BrandGarmentTestDoxList == null || !Req.BrandGarmentTestDoxList.Any())
                {
                    Req.BrandGarmentTestDoxList = new List<BrandGarmentTestDox>();
                }
                else
                {
                    foreach (var dox in Req.BrandGarmentTestDoxList.Where(o => o.IsOldFile == false))
                    {
                        dox.FilePath = this.UploadFilePath;
                        this.CreateMachineReportFile(dox);
                    }
                }

                _Provider.Processe_BrandGarmentTestDox(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model.Main = this.GetMain(new BrandGarmentTest_Request()
                {
                    ReportNo = Req.Main.ReportNo,
                }).Main;

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
        public BrandGarmentTest_ViewModel Delete(BrandGarmentTest_ViewModel Req)
        {
            BrandGarmentTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new BrandGarmentTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_BrandGarmentTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetMainList(new BrandGarmentTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new BrandGarmentTest_Request()
                {
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
        public BrandGarmentTest_ViewModel Download(BrandGarmentTest_Request Req)
        {
            BrandGarmentTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new BrandGarmentTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<BrandGarmentTest> tmpList = new List<BrandGarmentTest>();

                // 取得列表資料

                List<BrandGarmentTestDox> fileList = _Provider.GetBrandGarmentTestDoxList(Req);

                if (fileList.Any())
                {
                    // Zip檔若出現亂碼，須設定編碼方式為 UTF8
                    Ionic.Zip.ZipFile zipFile = new Ionic.Zip.ZipFile(encoding: Encoding.UTF8);
                    string zipName = $"BrandGarmentTestDox_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.zip";
                    string zipPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", zipName);

                    foreach (BrandGarmentTestDox dox in fileList)
                    {
                        string file = Path.Combine(dox.FilePath, dox.FileName);
                        zipFile.AddFile(file, string.Empty);
                    }


                    zipFile.Save(zipPath);
                    model.DownloadFileName = zipName;
                    model.DownloadFileFullName = zipPath;

                    model.Result = true;
                }
                else
                {
                    model.Result = false;
                    model.ErrorMessage = "File not found.";
                }

            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public BrandGarmentTest_ViewModel Download(List<BrandGarmentTestDox> ReqList)
        {
            BrandGarmentTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new BrandGarmentTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<BrandGarmentTest> tmpList = new List<BrandGarmentTest>();

                // 取得列表資料

                List<BrandGarmentTestDox> fileList = _Provider.GetBrandGarmentTestDoxList(ReqList);

                if (fileList.Any())
                {
                    // Zip檔若出現亂碼，須設定編碼方式為 UTF8
                    Ionic.Zip.ZipFile zipFile = new Ionic.Zip.ZipFile(encoding: Encoding.UTF8);
                    string zipName = $"BrandGarmentTestDox_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.zip";
                    string zipPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", zipName);

                    foreach (BrandGarmentTestDox dox in fileList)
                    {
                        string file = Path.Combine(dox.FilePath, dox.FileName);
                        zipFile.AddFile(file, string.Empty);
                    }


                    zipFile.Save(zipPath);
                    model.DownloadFileName = zipName;
                    model.DownloadFileFullName = zipPath;

                    model.Result = true;
                }
                else
                {
                    model.Result = false;
                    model.ErrorMessage = "File not found.";
                }

            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }
        public void CreateMachineReportFile(BrandGarmentTestDox dox)
        {
            dox.FileName = dox.BrandGarmentTestDoxFile.FileName;

            // 可能有各種檔案，因此需要取得副檔名
            string FileNameWithoutExtension = Path.GetFileNameWithoutExtension(dox.BrandGarmentTestDoxFile.FileName);
            string FileExtension = Path.GetExtension(dox.BrandGarmentTestDoxFile.FileName);

            if (!System.IO.Directory.Exists(this.UploadFilePath))
            {
                System.IO.Directory.CreateDirectory(this.UploadFilePath);
            }
            string FullFileName = $@"{this.UploadFilePath}\{dox.FileName}";

            string FileName = FileNameWithoutExtension;

            using (var fileStream = System.IO.File.Create(FullFileName))
            {
                dox.BrandGarmentTestDoxFile.InputStream.Seek(0, SeekOrigin.Begin);
                dox.BrandGarmentTestDoxFile.InputStream.CopyTo(fileStream);
            }
            dox.FileName = FileName + FileExtension;
        }
    }
}
