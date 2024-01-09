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
    public class BrandBulkTestService
    {
        private string UploadFileRootPath = ConfigurationManager.AppSettings["UploadFileRootPath"];
        private string UploadFilePath = $@"{ConfigurationManager.AppSettings["UploadFileRootPath"]}BulkFGT\BrandBulkTest\";

        private BrandBulkTestProvider _Provider;
        public BrandBulkTest_ViewModel GetDefaultModel(bool iNew = false)
        {
            BrandBulkTest_ViewModel model = new BrandBulkTest_ViewModel()
            {
                Request = new BrandBulkTest_Request(),
                Main = new BrandBulkTest()
                {
                    Result = "Pass"
                },
                MainList = new List<BrandBulkTest>(),
                BrandBulkTestDoxList = new List<BrandBulkTestDox>(),
                Article_Source = new List<SelectListItem>(),
            };

            try
            {
                _Provider = new BrandBulkTestProvider(Common.ProductionDataAccessLayer);
                model.Artwork_Source = _Provider.GetArtworkSource();

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
        public BrandBulkTest_ViewModel GetMainList(BrandBulkTest_Request Req)
        {
            BrandBulkTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new BrandBulkTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<BrandBulkTest> tmpList = new List<BrandBulkTest>();

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
        public BrandBulkTest_ViewModel GetMain(BrandBulkTest_Request Req)
        {
            BrandBulkTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new BrandBulkTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<BrandBulkTest> tmpList = new List<BrandBulkTest>();

                // 取得列表資料
                tmpList = _Provider.GetMainList(Req);


                if (tmpList.Any())
                {
                    model.Main = tmpList.FirstOrDefault();

                    List<BrandBulkTestDox> doxList = _Provider.GetBrandBulkTestDoxList(new BrandBulkTest_Request() { ReportNo = model.Main.ReportNo });

                    model.BrandBulkTestDoxList = doxList;

                    // 取得Article 下拉選單
                    _Provider = new BrandBulkTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new BrandBulkTest_Request() { OrderID = model.Main.OrderID });

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
        public BrandBulkTest_ViewModel GetOrderInfo(string OrderID)
        {
            BrandBulkTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new BrandBulkTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new BrandBulkTest_Request() { OrderID = OrderID });


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
        public BrandBulkTest_ViewModel NewSave(BrandBulkTest_ViewModel Req, string MDivision, string UserID)
        {
            BrandBulkTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new BrandBulkTestProvider(_ISQLDataTransaction);

                if (string.IsNullOrEmpty(Req.Main.OrderID) || string.IsNullOrEmpty(Req.Main.BrandID) || string.IsNullOrEmpty(Req.Main.SeasonID) || string.IsNullOrEmpty(Req.Main.StyleID) || string.IsNullOrEmpty(Req.Main.Article))
                {
                    throw new Exception("SP#、Brand、Season、Style and Article can't be empty.");
                }


                // 新增，並取得ReportNo
                _Provider.Insert_BrandBulkTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                // BrandBulkTestDoxList 檔案上傳
                if (Req.BrandBulkTestDoxList == null || !Req.BrandBulkTestDoxList.Any())
                {
                    Req.BrandBulkTestDoxList = new List<BrandBulkTestDox>();
                }
                else
                {
                    foreach (var dox in Req.BrandBulkTestDoxList.Where(o => o.IsOldFile == false))
                    {
                        dox.FilePath = this.UploadFilePath;
                        this.CreateMachineReportFile(dox);
                    }
                }


                _Provider.Processe_BrandBulkTestDox(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model.Main = this.GetMain(new BrandBulkTest_Request()
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
        public BrandBulkTest_ViewModel EditSave(BrandBulkTest_ViewModel Req, string UserID)
        {
            BrandBulkTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new BrandBulkTestProvider(_ISQLDataTransaction);


                // 更新表頭，並取得ReportNo
                _Provider.Update_BrandBulkTest(Req, UserID);


                // BrandBulkTestDoxList 檔案上傳
                if (Req.BrandBulkTestDoxList == null || !Req.BrandBulkTestDoxList.Any())
                {
                    Req.BrandBulkTestDoxList = new List<BrandBulkTestDox>();
                }
                else
                {
                    foreach (var dox in Req.BrandBulkTestDoxList.Where(o => o.IsOldFile == false))
                    {
                        dox.FilePath = this.UploadFilePath;
                        this.CreateMachineReportFile(dox);
                    }
                }

                _Provider.Processe_BrandBulkTestDox(Req, UserID);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model.Main = this.GetMain(new BrandBulkTest_Request()
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
        public BrandBulkTest_ViewModel Delete(BrandBulkTest_ViewModel Req)
        {
            BrandBulkTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new BrandBulkTestProvider(_ISQLDataTransaction);

                // 更新表頭，並取得ReportNo
                _Provider.Delete_BrandBulkTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetMainList(new BrandBulkTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                model.Request = new BrandBulkTest_Request()
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
        public BrandBulkTest_ViewModel Download(BrandBulkTest_Request Req)
        {
            BrandBulkTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new BrandBulkTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<BrandBulkTest> tmpList = new List<BrandBulkTest>();

                // 取得列表資料

                List<BrandBulkTestDox> fileList = _Provider.GetBrandBulkTestDoxList(Req);

                if (fileList.Any())
                {
                    // Zip檔若出現亂碼，須設定編碼方式為 UTF8
                    Ionic.Zip.ZipFile zipFile = new Ionic.Zip.ZipFile(encoding: Encoding.UTF8);
                    string zipName = $"BrandBulkTestDox_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.zip";
                    string zipPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", zipName);

                    foreach (BrandBulkTestDox dox in fileList)
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
        public BrandBulkTest_ViewModel Download(List<BrandBulkTestDox> ReqList)
        {
            BrandBulkTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new BrandBulkTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<BrandBulkTest> tmpList = new List<BrandBulkTest>();

                // 取得列表資料

                List<BrandBulkTestDox> fileList = _Provider.GetBrandBulkTestDoxList(ReqList);

                if (fileList.Any())
                {
                    // Zip檔若出現亂碼，須設定編碼方式為 UTF8
                    Ionic.Zip.ZipFile zipFile = new Ionic.Zip.ZipFile(encoding: Encoding.UTF8);
                    string zipName = $"BrandBulkTestDox_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.zip";
                    string zipPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", zipName);

                    foreach (BrandBulkTestDox dox in fileList)
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
        public void CreateMachineReportFile(BrandBulkTestDox dox)
        {
            dox.FileName = dox.BrandBulkTestDoxFile.FileName;

            // 可能有各種檔案，因此需要取得副檔名
            string FileNameWithoutExtension = Path.GetFileNameWithoutExtension(dox.BrandBulkTestDoxFile.FileName);
            string FileExtension = Path.GetExtension(dox.BrandBulkTestDoxFile.FileName);

            if (!System.IO.Directory.Exists(this.UploadFilePath))
            {
                System.IO.Directory.CreateDirectory(this.UploadFilePath);
            }
            string FullFileName = $@"{this.UploadFilePath}\{dox.FileName}";

            string FileName = FileNameWithoutExtension;

            using (var fileStream = System.IO.File.Create(FullFileName))
            {
                dox.BrandBulkTestDoxFile.InputStream.Seek(0, SeekOrigin.Begin);
                dox.BrandBulkTestDoxFile.InputStream.CopyTo(fileStream);
            }
            dox.FileName = FileName + FileExtension;
        }
    }
}
