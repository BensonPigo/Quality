using ADOHelper.Utility;
using DatabaseObject.ViewModel.BulkFGT;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Web.Mvc;
using System.Windows.Media.Animation;
using System.Web.WebPages;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System.Web.UI.WebControls;
using System.Web;
using Org.BouncyCastle.Asn1.Ocsp;
using static Sci.MyUtility;
using Microsoft.Office.Core;
using ClosedXML.Excel;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class EvaporationRateTestService
    {
        private EvaporationRateTestProvider _Provider;
        private MailToolsService _MailService;
        public EvaporationRateTest_ViewModel GetDefaultModel(bool IsNew = false)
        {
            EvaporationRateTest_ViewModel model = new EvaporationRateTest_ViewModel()
            {
                Request = new EvaporationRateTest_Request(),
                Main = new EvaporationRateTest_Main()
                {
                    // ISP20230288 寫死4
                    //MachineReport = "",
                },

                ReportNo_Source = new List<System.Web.Mvc.SelectListItem>(),
                DetailList = new List<EvaporationRateTest_Detail>(),
                SpecimenList = new List<EvaporationRateTest_Specimen>(),
                TimeList = new List<EvaporationRateTest_Specimen_Time>(),
                //RateList = new List<EvaporationRateTest_Specimen_Rate>(),
                Article_Source = new List<SelectListItem>(),
            };

            try
            {

                if (IsNew)
                {
                    //Before
                    model.DetailList.Add(new EvaporationRateTest_Detail()
                    {
                        Type = "Before"
                    });
                    //After
                    model.DetailList.Add(new EvaporationRateTest_Detail()
                    {
                        Type = "After"
                    });

                    // EvaporationRateTest_Specimen
                    foreach (var item in model.DetailList)
                    {
                        string type = item.Type;
                        model.SpecimenList.Add(new EvaporationRateTest_Specimen()
                        {
                            DetailType = type,
                            SpecimenID = "Specimen 1"
                        });
                        model.SpecimenList.Add(new EvaporationRateTest_Specimen()
                        {
                            DetailType = type,
                            SpecimenID = "Specimen 2"
                        });
                        model.SpecimenList.Add(new EvaporationRateTest_Specimen()
                        {
                            DetailType = type,
                            SpecimenID = "Specimen 3"
                        });
                        model.SpecimenList.Add(new EvaporationRateTest_Specimen()
                        {
                            DetailType = type,
                            SpecimenID = "Specimen 4"
                        });
                        model.SpecimenList.Add(new EvaporationRateTest_Specimen()
                        {
                            DetailType = type,
                            SpecimenID = "Specimen 5"
                        });
                    }

                    // EvaporationRateTest_Specimen_Time
                    foreach (var specimen in model.SpecimenList)
                    {
                        model.TimeList.Add(new EvaporationRateTest_Specimen_Time()
                        {
                            DetailType = specimen.DetailType,
                            SpecimenID = specimen.SpecimenID,
                            Time = 0,
                            IsInitialMass = true,
                        });

                        //for (int time = 0; time <= 15; time += 3)
                        //{
                        //    model.TimeList.Add(new EvaporationRateTest_Specimen_Time()
                        //    {
                        //        DetailType = specimen.DetailType,
                        //        SpecimenID = specimen.SpecimenID,
                        //        Time = time,
                        //        IsInitialMass = time == 0,
                        //    });
                        //}
                    }
                }
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
        /// 取得新的表身資料組合
        /// </summary>
        /// <param name="SpecimenUkey"></param>
        /// <param name="lastTime"></param>
        /// <param name="inputIsInitialTime"></param>
        /// <param name="inputIsInitialSpecimen_TimeUkey"></param>
        /// <param name="lastRateName"></param>
        /// <param name="lastMinuend_Time"></param>
        /// <returns></returns>
        public EvaporationRateTest_ViewModel GetNewTimeRate(long SpecimenUkey, int lastTime, int inputIsInitialTime, long inputIsInitialSpecimen_TimeUkey, string lastRateName, int lastMinuend_Time)
        {
            EvaporationRateTest_ViewModel model = new EvaporationRateTest_ViewModel()
            {
                TimeList = new List<EvaporationRateTest_Specimen_Time>(),
                RateList = new List<EvaporationRateTest_Specimen_Rate>(),
            };

            try
            {
                model.TimeList.Add(new EvaporationRateTest_Specimen_Time()
                {
                    SpecimenUkey = SpecimenUkey,
                    Time = lastTime + 10,
                    IsInitialMass = false,
                    InitialTime = inputIsInitialTime, // Time = 0 
                    InitialTimeUkey = inputIsInitialSpecimen_TimeUkey // Time = 0 的Ukey
                });
                //
                model.RateList.Add(new EvaporationRateTest_Specimen_Rate()
                {
                    SpecimenUkey = SpecimenUkey,
                    RateName = "R" + lastRateName[1] + 1,
                    Subtrahend_Time = lastMinuend_Time, //前一個的被減數 = 下一個的減數
                    Minuend_Time = lastMinuend_Time + 10,
                });

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
        public EvaporationRateTest_ViewModel GetData(EvaporationRateTest_Request Req)
        {
            EvaporationRateTest_ViewModel model = this.GetDefaultModel();

            try
            {
                _Provider = new EvaporationRateTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<EvaporationRateTest_Main> tmpList = new List<EvaporationRateTest_Main>();

                if (string.IsNullOrEmpty(Req.BrandID) && string.IsNullOrEmpty(Req.SeasonID) && string.IsNullOrEmpty(Req.StyleID) && string.IsNullOrEmpty(Req.Article) && !string.IsNullOrEmpty(Req.ReportNo))
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new EvaporationRateTest_Request()
                    {
                        ReportNo = Req.ReportNo,
                    });
                }
                else
                {
                    // 根據四大天王，取得符合條件的主表
                    tmpList = _Provider.GetMainList(new EvaporationRateTest_Request()
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

                    model.Request = Req;
                    model.Request.ReportNo = model.Main.ReportNo;

                    // 取得表身資料
                    model.DetailList = _Provider.GetDetailList(new EvaporationRateTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    // 取得表身資料
                    model.SpecimenList = _Provider.GetSpecimenList(new EvaporationRateTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });
                    model.TimeList = _Provider.GetTimeList(new EvaporationRateTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });

                    //model.RateList = _Provider.GetRateList(new EvaporationRateTest_Request()
                    //{
                    //    ReportNo = model.Main.ReportNo
                    //});



                    // 取得Article 下拉選單
                    _Provider = new EvaporationRateTestProvider(Common.ProductionDataAccessLayer);
                    List<DatabaseObject.ProductionDB.Orders> tmpOrders = _Provider.GetOrderInfo(new EvaporationRateTest_Request() { OrderID = model.Main.OrderID });

                    foreach (var oriData in tmpOrders)
                    {
                        SelectListItem Article = new SelectListItem()
                        {
                            Text = oriData.Article,
                            Value = oriData.Article,
                        };
                        model.Article_Source.Add(Article);
                    }

                    string Subject = $"Evaporation Rate Test/{model.Main.OrderID}/" +
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

        public EvaporationRateTest_ViewModel GetOrderInfo(string OrderID)
        {
            EvaporationRateTest_ViewModel model = this.GetDefaultModel();
            List<DatabaseObject.ProductionDB.Orders> tmpOrders = new List<DatabaseObject.ProductionDB.Orders>();
            try
            {
                _Provider = new EvaporationRateTestProvider(Common.ProductionDataAccessLayer);

                tmpOrders = _Provider.GetOrderInfo(new EvaporationRateTest_Request() { OrderID = OrderID });


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
        public EvaporationRateTest_ViewModel NewSave(EvaporationRateTest_ViewModel Req, string MDivision, string UserID)
        {
            EvaporationRateTest_ViewModel model = this.GetDefaultModel();
            string newReportNo = string.Empty;
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new EvaporationRateTestProvider(_ISQLDataTransaction);

                if (Req.DetailList == null || !Req.DetailList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Detail can not be empty.";
                    return model;
                }
                if (Req.SpecimenList == null || !Req.SpecimenList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Specimen List can not be empty.";
                    return model;
                }
                if (Req.TimeList == null || !Req.TimeList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Time List can not be empty.";
                    return model;
                }
                //if (Req.RateList == null || !Req.RateList.Any())
                //{
                //    model.Result = false;
                //    model.ErrorMessage = "Rate List can not be empty.";
                //    return model;
                //}

                // 新增，並取得ReportNo
                _Provider.Insert_EvaporationRateTest(Req, MDivision, UserID, out newReportNo);
                Req.Main.ReportNo = newReportNo;

                _Provider.Processe_EvaporationRateTest_Detail(Req, UserID, out List<EvaporationRateTest_Detail> NewDetailList);

                // 若有Detail，則把Before After的Type寫進下一層
                if (NewDetailList.Any())
                {
                    foreach (var newDetail in NewDetailList)
                    {
                        string type = newDetail.Type;
                        foreach (var specimen in Req.SpecimenList.Where(o => o.DetailType == type))
                        {
                            specimen.DetailUkey = newDetail.Ukey;
                        }
                    }
                }

                _Provider.Processe_EvaporationRateTest_Specimen(Req, UserID, out List<EvaporationRateTest_Specimen> NewSpecimenList);
                // 若有Specimen，則把SpecimenUkey寫進下一層
                if (NewSpecimenList.Any())
                {
                    foreach (var newSpecimenin in NewSpecimenList.Distinct().ToList())
                    {
                        foreach (var time in Req.TimeList.Where(o => o.DetailType == newSpecimenin.DetailType && o.SpecimenID == newSpecimenin.SpecimenID))
                        {
                            time.SpecimenUkey = newSpecimenin.Ukey;
                        }

                        //foreach (var rate in Req.RateList.Where(o => o.DetailType == newSpecimenin.DetailType && o.SpecimenID == newSpecimenin.SpecimenID))
                        //{
                        //    rate.SpecimenUkey = newSpecimenin.Ukey;
                        //}

                    }
                }

                _Provider.Processe_EvaporationRateTest_Specimen_Time(Req, UserID, out List<EvaporationRateTest_Specimen_Time> NewTimeList);
                //if (NewTimeList.Any())
                //{

                //    int idx = 0;
                //    foreach (var rate in Req.RateList)
                //    {
                //        // R1 = Time 0 - Time 10
                //        // R2 = Time 10 - Time 20...以此類推
                //        // 但最後兩個不相減
                //        var Subtrahend = NewTimeList.Where(o => o.DetailType == rate.DetailType && o.SpecimenID == rate.SpecimenID && o.Time == rate.Subtrahend_Time).FirstOrDefault();
                //        var Minuend = NewTimeList.Where(o => o.DetailType == rate.DetailType && o.SpecimenID == rate.SpecimenID && o.Time == rate.Minuend_Time).FirstOrDefault();

                //        rate.Subtrahend_TimeUkey = Subtrahend.Ukey;
                //        rate.Minuend_TimeUkey = Minuend.Ukey;
                //    }
                //}
                //_Provider.Processe_EvaporationRateTest_Specimen_Rate(Req, UserID);

                _ISQLDataTransaction.Commit();

                //_Provider.UpdateAverage(Req.Main);

                // 重新查詢資料
                model = this.GetData(new EvaporationRateTest_Request()
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
        public EvaporationRateTest_ViewModel EditSave(EvaporationRateTest_ViewModel Req, string UserID)
        {
            EvaporationRateTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new EvaporationRateTestProvider(_ISQLDataTransaction);

                if (Req.DetailList == null || !Req.DetailList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Detail can not be empty.";
                    return model;
                }
                if (Req.SpecimenList == null || !Req.SpecimenList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Specimen List can not be empty.";
                    return model;
                }
                if (Req.TimeList == null || !Req.TimeList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Time List can not be empty.";
                    return model;
                }
                //if (Req.RateList == null || !Req.RateList.Any())
                //{
                //    model.Result = false;
                //    model.ErrorMessage = "Rate List can not be empty.";
                //    return model;
                //}

                // 更新表頭，並取得ReportNo
                _Provider.Update_EvaporationRateTest(Req, UserID);

                // 更新表身

                _Provider.Processe_EvaporationRateTest_Detail(Req, UserID, out List<EvaporationRateTest_Detail> NewDetailList);

                _Provider.Processe_EvaporationRateTest_Specimen(Req, UserID, out List<EvaporationRateTest_Specimen> NewSpecimenList);
                // 若有Specimen，則把SpecimenUkey寫進下一層
                if (NewSpecimenList.Any())
                {
                    foreach (var newSpecimenin in NewSpecimenList.Distinct().ToList())
                    {
                        foreach (var time in Req.TimeList.Where(o => o.DetailType == newSpecimenin.DetailType && o.SpecimenID == newSpecimenin.SpecimenID))
                        {
                            time.SpecimenUkey = newSpecimenin.Ukey;
                        }

                        //foreach (var rate in Req.RateList.Where(o => o.DetailType == newSpecimenin.DetailType && o.SpecimenID == newSpecimenin.SpecimenID))
                        //{
                        //    rate.SpecimenUkey = newSpecimenin.Ukey;
                        //}

                    }
                }
                _Provider.Processe_EvaporationRateTest_Specimen_Time(Req, UserID, out List<EvaporationRateTest_Specimen_Time> NewTimeList);

                //if (NewTimeList.Any())
                //{

                //    int idx = 0;
                //    foreach (var rate in Req.RateList)
                //    {
                //        // R1 = Time 0 - Time 10
                //        // R2 = Time 10 - Time 20...以此類推
                //        // 但最後兩個不相減
                //        var Subtrahend = NewTimeList.Where(o => o.DetailType == rate.DetailType && o.SpecimenID == rate.SpecimenID && o.Time == rate.Subtrahend_Time).FirstOrDefault();
                //        var Minuend = NewTimeList.Where(o => o.DetailType == rate.DetailType && o.SpecimenID == rate.SpecimenID && o.Time == rate.Minuend_Time).FirstOrDefault();

                //        rate.Subtrahend_TimeUkey = Subtrahend.Ukey;
                //        rate.Minuend_TimeUkey = Minuend.Ukey;
                //    }
                //}
                //_Provider.Processe_EvaporationRateTest_Specimen_Rate(Req, UserID);

                //_Provider.UpdateAverage(Req.Main);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new EvaporationRateTest_Request()
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
        public EvaporationRateTest_ViewModel Delete(EvaporationRateTest_ViewModel Req)
        {
            EvaporationRateTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);


            try
            {
                _Provider = new EvaporationRateTestProvider(_ISQLDataTransaction);


                // 更新表頭，並取得ReportNo
                _Provider.Delete_EvaporationRateTest(Req);

                _ISQLDataTransaction.Commit();

                // 重新查詢資料
                model = this.GetData(new EvaporationRateTest_Request()
                {
                    BrandID = Req.Main.BrandID,
                    SeasonID = Req.Main.SeasonID,
                    StyleID = Req.Main.StyleID,
                    Article = Req.Main.Article,
                });

                if (!string.IsNullOrEmpty(model.Main.ReportNo))
                {
                    model.Request = new EvaporationRateTest_Request()
                    {
                        ReportNo = model.Main.ReportNo,
                        BrandID = model.Main.BrandID,
                        SeasonID = model.Main.SeasonID,
                        StyleID = model.Main.StyleID,
                        Article = model.Main.Article,
                    };
                }
                else
                {
                    model.Result = true;
                }
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
        public EvaporationRateTest_ViewModel EncodeAmend(EvaporationRateTest_ViewModel Req, string UserID)
        {
            EvaporationRateTest_ViewModel model = this.GetDefaultModel();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ManufacturingExecutionDataAccessLayer);

            try
            {
                _Provider = new EvaporationRateTestProvider(_ISQLDataTransaction);


                // 更新表頭，並取得ReportNo
                _Provider.EncodeAmend_EvaporationRateTest(Req.Main, UserID);

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
        public EvaporationRateTest_ViewModel GetReport(string ReportNo, bool isPDF, string AssignedFineName = "")
        {
            EvaporationRateTest_ViewModel result = new EvaporationRateTest_ViewModel();

            string basefileName = "EvaporationRateTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";
            string tmpName = string.Empty;
            Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);

            try
            {

                // 取得報表資料

                EvaporationRateTest_ViewModel model = this.GetData(new EvaporationRateTest_Request() { ReportNo = ReportNo });

                _Provider = new EvaporationRateTestProvider(Common.ManufacturingExecutionDataAccessLayer);
                //model.Item_Source = _Provider.GetTestItems();

                DataTable ReportTechnician = _Provider.GetReportTechnician(new EvaporationRateTest_Request() { ReportNo = ReportNo });
                excel.Visible = false;
                excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
                Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表

                // 取得工作表上所有圖形物件
                Microsoft.Office.Interop.Excel.Shapes shapes = worksheet.Shapes;

                // 根據名稱，搜尋文字方塊物件
                Microsoft.Office.Interop.Excel.Shape beforeChart1 = shapes.Item("BeforeChart1");
                Microsoft.Office.Interop.Excel.Shape beforeChart2 = shapes.Item("BeforeChart2");
                Microsoft.Office.Interop.Excel.Shape beforeChart3 = shapes.Item("BeforeChart3");
                Microsoft.Office.Interop.Excel.Shape beforeChart4 = shapes.Item("BeforeChart4");
                Microsoft.Office.Interop.Excel.Shape beforeChart5 = shapes.Item("BeforeChart5");

                tmpName = $"Evaporation Rate Test_{model.Main.OrderID}_" +
                   $"{model.Main.StyleID}_" +
                   $"{model.Main.FabricRefNo}_" +
                   $"{model.Main.FabricColor}_" +
                   $"{model.Main.Result}_" +
                   $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                #region Before Sheet

                worksheet.Cells[3, 10] = model.Main.ReportNo;
                worksheet.Cells[3, 17] = model.Main.ReportDateText;
                worksheet.Cells[3, 24] = model.Main.SubmitDateText;

                worksheet.Cells[4, 10] = model.Main.OrderID;
                worksheet.Cells[4, 17] = model.Main.BrandID;
                worksheet.Cells[4, 24] = model.Main.SeasonID;

                worksheet.Cells[5, 10] = model.Main.StyleID;
                worksheet.Cells[5, 17] = model.Main.Article;
                worksheet.Cells[5, 24] = model.Main.Seq;

                worksheet.Cells[6, 10] = model.Main.FabricRefNo;
                worksheet.Cells[6, 17] = model.Main.FabricColor;
                worksheet.Cells[6, 24] = model.Main.FabricDescription;

                // Technician 、 Approver 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["TechnicianName"].ToString();

                    // 姓名
                    worksheet.Cells[63, 9] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[65, 9];
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

                    string ApproverName = ReportTechnician.Rows[0]["ApproverName"].ToString();

                    // 姓名
                    worksheet.Cells[63, 23] = ApproverName;

                    // Signture 圖片
                    cell = worksheet.Cells[65, 23];
                    if (ReportTechnician.Rows[0]["ApproverSignture"] != DBNull.Value)
                    {

                        byte[] TestBeforePicture = (byte[])ReportTechnician.Rows[0]["ApproverSignture"]; // 圖片的 byte[]

                        MemoryStream ms = new MemoryStream(TestBeforePicture);
                        System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                        string imageName = $"{Guid.NewGuid()}.jpg";
                        string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                        img.Save(imgPath);
                        worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    }
                }

                var specimenBeforeList = model.SpecimenList.Where(o => o.DetailType == "Before").OrderBy(o => o.SpecimenID).ToList();

                Dictionary<string, int> dicSpecimenColIndex = new Dictionary<string, int>()
                {
                    ["Specimen 1"] = 2,
                    ["Specimen 2"] = 9,
                    ["Specimen 3"] = 16,
                    ["Specimen 4"] = 23,
                    ["Specimen 5"] = 30,
                };

                int specimenStartColIdx = 2;
                int specimenAvgIdx = 53;

                // specimen逐一開始填入Mass
                foreach (var specimen in specimenBeforeList)
                {
                    // B15
                    int massInputRowIndex = 15;
                    int massInputColIdx = dicSpecimenColIndex[specimen.SpecimenID];
                    int ctnTimeRow = 0;

                    var timeListt = model.TimeList.Where(o => o.DetailType == "Before" && o.SpecimenID == specimen.SpecimenID).OrderBy(o => o.Time).ToList();

                    var specimenAvg = specimen.RateAverage;
                    worksheet.Cells[specimenAvgIdx, 12] = specimenAvg; // L53
                    if (specimenAvg == 0)
                    {
                        worksheet.Cells[specimenAvgIdx, 12] = null;
                        worksheet.Cells[specimenAvgIdx, 13] = null;
                    }
                    specimenAvgIdx++;

                    // 若整個Mass都為空，把不需要的公式清空，跳過Specimen
                    if (!timeListt.Any(o => o.Mass > 0))
                    {
                        continue;
                    }

                    // Mss填入
                    foreach (var t in timeListt)
                    {
                        if (t.IsInitialMass)
                        {
                            //worksheet.Cells[13, massInputColIdx] = t.Mass; // B13
                            worksheet.Cells[13, massInputColIdx + 2] = t.Mass; // D13
                        }
                        worksheet.Cells[massInputRowIndex, massInputColIdx] = t.Mass;
                        massInputRowIndex++;
                        ctnTimeRow++;
                    }

                    // 清空剩下的公式，等於20代表剛好填滿60分鐘
                    if (ctnTimeRow < 20)
                    {
                        // B21:F35
                        Range range = worksheet.Range[$"{MyUtility.Excel.ConvertNumericToExcelColumn(specimenStartColIdx)}{massInputRowIndex}:{MyUtility.Excel.ConvertNumericToExcelColumn(specimenStartColIdx + 4)}35"];
                        // 清空範圍的值
                        range.ClearContents();
                    }

                    // 設定圖表的資料範圍，由於前面欄位是多抓一格，所以這邊扣掉1才是正確的範圍
                    massInputRowIndex--;
                    if (specimen.SpecimenID == "Specimen 1")
                    {
                        Chart chart = beforeChart1.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='BEFORE WASH'!$A$15:$A${massInputRowIndex},'BEFORE WASH'!$C$15:$C${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);

                    }
                    else if (specimen.SpecimenID == "Specimen 2")
                    {
                        Chart chart = beforeChart2.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='BEFORE WASH'!$H$15:$H${massInputRowIndex},'BEFORE WASH'!$J$15:$J${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);
                    }
                    else if (specimen.SpecimenID == "Specimen 3")
                    {
                        Chart chart = beforeChart3.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='BEFORE WASH'!$O$15:$O${massInputRowIndex},'BEFORE WASH'!$Q$15:$Q${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);
                    }
                    else if (specimen.SpecimenID == "Specimen 4")
                    {
                        Chart chart = beforeChart4.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='BEFORE WASH'!$V$15:$V${massInputRowIndex},'BEFORE WASH'!$X$15:$X${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);
                    }
                    else if (specimen.SpecimenID == "Specimen 5")
                    {
                        Chart chart = beforeChart5.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='BEFORE WASH'!$AC$15:$AC${massInputRowIndex},'BEFORE WASH'!$AE$15:$AE${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);
                    }


                    // 下一個Specimen間隔6 column
                    specimenStartColIdx += 7;
                }

                #endregion

                specimenStartColIdx = 2;
                #region After Sheet
                worksheet = excel.ActiveWorkbook.Worksheets[2];
                shapes = worksheet.Shapes;
                Microsoft.Office.Interop.Excel.Shape afterChart1 = shapes.Item("AfterChart1");
                Microsoft.Office.Interop.Excel.Shape afterChart2 = shapes.Item("AfterChart2");
                Microsoft.Office.Interop.Excel.Shape afterChart3 = shapes.Item("AfterChart3");
                Microsoft.Office.Interop.Excel.Shape afterChart4 = shapes.Item("AfterChart4");
                Microsoft.Office.Interop.Excel.Shape afterChart5 = shapes.Item("AfterChart5");

                worksheet.Cells[3, 10] = model.Main.ReportNo;
                worksheet.Cells[3, 17] = model.Main.ReportDateText;
                worksheet.Cells[3, 24] = model.Main.SubmitDateText;

                worksheet.Cells[4, 10] = model.Main.OrderID;
                worksheet.Cells[4, 17] = model.Main.BrandID;
                worksheet.Cells[4, 24] = model.Main.SeasonID;

                worksheet.Cells[5, 10] = model.Main.StyleID;
                worksheet.Cells[5, 17] = model.Main.Article;
                worksheet.Cells[5, 24] = model.Main.Seq;

                worksheet.Cells[6, 10] = model.Main.FabricRefNo;
                worksheet.Cells[6, 17] = model.Main.FabricColor;
                worksheet.Cells[6, 24] = model.Main.FabricDescription;

                // Technician 、 Approver 欄位
                if (ReportTechnician.Rows != null && ReportTechnician.Rows.Count > 0)
                {
                    string TechnicianName = ReportTechnician.Rows[0]["TechnicianName"].ToString();

                    // 姓名
                    worksheet.Cells[63, 9] = TechnicianName;

                    // Signture 圖片
                    Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[65, 9];
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

                    string ApproverName = ReportTechnician.Rows[0]["ApproverName"].ToString();

                    // 姓名
                    worksheet.Cells[63, 23] = ApproverName;

                    // Signture 圖片
                    cell = worksheet.Cells[65, 23];
                    if (ReportTechnician.Rows[0]["ApproverSignture"] != DBNull.Value)
                    {

                        byte[] TestBeforePicture = (byte[])ReportTechnician.Rows[0]["ApproverSignture"]; // 圖片的 byte[]

                        MemoryStream ms = new MemoryStream(TestBeforePicture);
                        System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                        string imageName = $"{Guid.NewGuid()}.jpg";
                        string imgPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", imageName);

                        img.Save(imgPath);
                        worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cell.Left, cell.Top, 100, 24);

                    }
                }

                var specimenAfterList = model.SpecimenList.Where(o => o.DetailType == "After").OrderBy(o => o.SpecimenID).ToList();

                specimenStartColIdx = 2;
                specimenAvgIdx = 53;
                // 開始填入Mass
                foreach (var specimen in specimenAfterList)
                {
                    // B15
                    int massInputRowIndex = 15;
                    int massInputColIdx = dicSpecimenColIndex[specimen.SpecimenID];
                    int ctnTimeRow = 0;

                    var timeListt = model.TimeList.Where(o => o.DetailType == "After" && o.SpecimenID == specimen.SpecimenID).OrderBy(o => o.Time).ToList();

                    var specimenAvg = specimen.RateAverage;
                    worksheet.Cells[specimenAvgIdx, 12] = specimenAvg; // L53
                    if (specimenAvg == 0)
                    {
                        worksheet.Cells[specimenAvgIdx, 12] = null;
                        worksheet.Cells[specimenAvgIdx, 13] = null;
                    }
                    specimenAvgIdx++;

                    // 若整個Mass都為空，把不需要的公式清空，跳過Specimen
                    if (!timeListt.Any(o => o.Mass > 0))
                    {
                        continue;
                    }

                    // Mss填入
                    foreach (var t in timeListt)
                    {
                        if (t.IsInitialMass)
                        {
                            //worksheet.Cells[13, massInputColIdx] = t.Mass; // B13
                            worksheet.Cells[13, massInputColIdx + 2] = t.Mass; // D13
                        }
                        worksheet.Cells[massInputRowIndex, massInputColIdx] = t.Mass;
                        massInputRowIndex++;
                        ctnTimeRow++;
                    }

                    // 清空剩下的公式，等於20代表剛好填滿60分鐘
                    if (ctnTimeRow < 20)
                    {
                        // B21:F35
                        Range range = worksheet.Range[$"{MyUtility.Excel.ConvertNumericToExcelColumn(specimenStartColIdx)}{massInputRowIndex}:{MyUtility.Excel.ConvertNumericToExcelColumn(specimenStartColIdx + 4)}35"];
                        // 清空範圍的值
                        range.ClearContents();
                    }

                    // 設定圖表的資料範圍，由於前面欄位是多抓一格，所以這邊扣掉1才是正確的範圍
                    massInputRowIndex--;
                    if (specimen.SpecimenID == "Specimen 1")
                    {
                        Chart chart = afterChart1.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='AFTER 5TH WASH'!$A$15:$A${massInputRowIndex},'AFTER 5TH WASH'!$C$15:$C${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);

                    }
                    else if (specimen.SpecimenID == "Specimen 2")
                    {
                        Chart chart = afterChart2.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='AFTER 5TH WASH'!$H$15:$H${massInputRowIndex},'AFTER 5TH WASH'!$J$15:$J${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);
                    }
                    else if (specimen.SpecimenID == "Specimen 3")
                    {
                        Chart chart = afterChart3.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='AFTER 5TH WASH'!$O$15:$O${massInputRowIndex},'AFTER 5TH WASH'!$Q$15:$Q${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);
                    }
                    else if (specimen.SpecimenID == "Specimen 4")
                    {
                        Chart chart = afterChart4.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='AFTER 5TH WASH'!$V$15:$V${massInputRowIndex},'AFTER 5TH WASH'!$X$15:$X${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);
                    }
                    else if (specimen.SpecimenID == "Specimen 5")
                    {
                        Chart chart = afterChart5.Chart;

                        // 設定新的資料範圍，還沒完成
                        Range newDataRange = worksheet.Range[$@"='AFTER 5TH WASH'!$AC$15:$AC${massInputRowIndex},'AFTER 5TH WASH'!$AE$15:$AE${massInputRowIndex}"]; // 替換為你的資料範圍
                        chart.SetSourceData(newDataRange);
                    }

                    // 下一個Specimen間隔6 column
                    specimenStartColIdx += 7;
                }

                #endregion

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }
                char[] invalidChars = Path.GetInvalidFileNameChars();
                char[] additionalChars = { '-', '+' }; // 您想要新增的字元
                char[] updatedInvalidChars = invalidChars.Concat(additionalChars).ToArray();

                foreach (char invalidChar in updatedInvalidChars)
                {
                    tmpName = tmpName.Replace(invalidChar.ToString(), "");
                }

                #region Title
                worksheet = excel.ActiveWorkbook.Worksheets[1];

                string FactoryNameEN = _Provider.GetFactoryNameEN(ReportNo, System.Web.HttpContext.Current.Session["FactoryID"].ToString());

                // 1. 插入一列
                worksheet.Rows["1"].Insert();
                // 2. 合併欄位 (B1:K1)
                Microsoft.Office.Interop.Excel.Range mergedRange = worksheet.Range["A1", "AH1"];
                mergedRange.Merge();

                // 設置字體樣式

                // 3. 設置文字和樣式
                mergedRange.Value = FactoryNameEN; // 替換為你的 FactoryNameEN 變數
                mergedRange.Font.Name = "Arial";      // 設置字體類型
                mergedRange.Font.Size = 25;          // 設置字體大小
                mergedRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // 水平置中
                mergedRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;   // 垂直置中
                mergedRange.Font.Bold = true;        // 設置字體加粗

                // 自動檢測使用範圍

                Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
                if (usedRange != null)
                {
                    // 獲取最後一行
                    int lastRow = usedRange.Row + usedRange.Rows.Count - 1;

                    // 清除已有的列印範圍
                    worksheet.PageSetup.PrintArea = string.Empty;

                    // 設定新的列印範圍
                    string printArea = $"A1:AH{lastRow + 10}";
                    worksheet.PageSetup.PrintArea = printArea;
                }
                #endregion

                #region Title
                worksheet = excel.ActiveWorkbook.Worksheets[2];
                // 1. 插入一列
                worksheet.Rows["1"].Insert();
                // 2. 合併欄位 (B1:K1)
                Microsoft.Office.Interop.Excel.Range mergedRange1 = worksheet.Range["A1", "AH1"];
                mergedRange1.Merge();

                // 設置字體樣式

                // 3. 設置文字和樣式
                mergedRange1.Value = FactoryNameEN; // 替換為你的 FactoryNameEN 變數
                mergedRange1.Font.Name = "Arial";      // 設置字體類型
                mergedRange1.Font.Size = 25;          // 設置字體大小
                mergedRange1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; // 水平置中
                mergedRange1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;   // 垂直置中
                mergedRange1.Font.Bold = true;        // 設置字體加粗

                // 自動檢測使用範圍

                Microsoft.Office.Interop.Excel.Range usedRange1 = worksheet.UsedRange;
                if (usedRange1 != null)
                {
                    // 獲取最後一行
                    int lastRow = usedRange1.Row + usedRange1.Rows.Count - 1;

                    // 清除已有的列印範圍
                    worksheet.PageSetup.PrintArea = string.Empty;

                    // 設定新的列印範圍
                    string printArea = $"A1:AH{lastRow + 10}";
                    worksheet.PageSetup.PrintArea = printArea;
                }
                #endregion

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
                    LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                    officeService.ConvertExcelToPdf(fullExcelFileName, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));

                    result.TempFileName = filePdfName;
                }
                else
                {
                    result.TempFileName = fileName;
                }
                result.Result = true;
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
        public string ConvertNumberToLetter(int number)
        {
            if (number >= 1 && number <= 26)
            {
                char letter = (char)('A' + number - 1);
                return letter.ToString();
            }
            else
            {
                return "Invalid";
            }
        }

        public SendMail_Result SendMail(string ReportNo, string TO, string CC, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            _Provider = new EvaporationRateTestProvider(Common.ManufacturingExecutionDataAccessLayer);

            EvaporationRateTest_ViewModel model = this.GetData(new EvaporationRateTest_Request() { ReportNo = ReportNo });

            string name = $"Evaporation Rate Test_{model.Main.OrderID}_" +
                    $"{model.Main.StyleID}_" +
                    $"{model.Main.FabricRefNo}_" +
                    $"{model.Main.FabricColor}_" +
                    $"{model.Main.Result}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
            EvaporationRateTest_ViewModel report = this.GetReport(ReportNo, false, name);
            string mailBody = "";
            string FileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", report.TempFileName);
            SendMail_Request sendMail_Request = new SendMail_Request
            {
                Subject = $"Evaporation Rate Test/{model.Main.OrderID}/" +
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
