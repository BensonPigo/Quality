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

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class EvaporationRateTestService
    {
        private EvaporationRateTestProvider _Provider;
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
                RateList = new List<EvaporationRateTest_Specimen_Rate>(),
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
                        model.TimeList.Add(new EvaporationRateTest_Specimen_Time()
                        {
                            DetailType = specimen.DetailType,
                            SpecimenID = specimen.SpecimenID,
                            Time = 10,
                            IsInitialMass = false,
                            InitialTime = 0
                        });
                        model.TimeList.Add(new EvaporationRateTest_Specimen_Time()
                        {
                            DetailType = specimen.DetailType,
                            SpecimenID = specimen.SpecimenID,
                            Time = 20,
                            IsInitialMass = false,
                            InitialTime = 0
                        });
                        model.TimeList.Add(new EvaporationRateTest_Specimen_Time()
                        {
                            DetailType = specimen.DetailType,
                            SpecimenID = specimen.SpecimenID,
                            Time = 30,
                            IsInitialMass = false,
                            InitialTime = 0
                        });
                        model.TimeList.Add(new EvaporationRateTest_Specimen_Time()
                        {
                            DetailType = specimen.DetailType,
                            SpecimenID = specimen.SpecimenID,
                            Time = 40,
                            IsInitialMass = false,
                            InitialTime = 0
                        });
                    }

                    // EvaporationRateTest_Specimen_Rate
                    foreach (var specimen in model.SpecimenList)
                    {
                        model.RateList.Add(new EvaporationRateTest_Specimen_Rate()
                        {
                            DetailType = specimen.DetailType,
                            SpecimenID = specimen.SpecimenID,
                            RateName = "R1",
                            Subtrahend_Time = 0,
                            Minuend_Time = 10,
                            Ratio = 6,
                        });
                        model.RateList.Add(new EvaporationRateTest_Specimen_Rate()
                        {
                            DetailType = specimen.DetailType,
                            SpecimenID = specimen.SpecimenID,
                            RateName = "R2",
                            Subtrahend_Time = 10,
                            Minuend_Time = 20,
                            Ratio = 6,
                        });
                        model.RateList.Add(new EvaporationRateTest_Specimen_Rate()
                        {
                            DetailType = specimen.DetailType,
                            SpecimenID = specimen.SpecimenID,
                            RateName = "R3",
                            Subtrahend_Time = 20,
                            Minuend_Time = 30,
                            Ratio = 6,
                        });
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

                    model.RateList = _Provider.GetRateList(new EvaporationRateTest_Request()
                    {
                        ReportNo = model.Main.ReportNo
                    });



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

                if (!Req.DetailList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Detail can not be empty.";
                    return model;
                }
                if (!Req.SpecimenList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Specimen List can not be empty.";
                    return model;
                }
                if (!Req.TimeList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Time List can not be empty.";
                    return model;
                }
                if (!Req.RateList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Rate List can not be empty.";
                    return model;
                }

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

                        foreach (var rate in Req.RateList.Where(o => o.DetailType == newSpecimenin.DetailType && o.SpecimenID == newSpecimenin.SpecimenID))
                        {
                            rate.SpecimenUkey = newSpecimenin.Ukey;
                        }

                    }
                }

                _Provider.Processe_EvaporationRateTest_Specimen_Time(Req, UserID, out List<EvaporationRateTest_Specimen_Time> NewTimeList);
                if (NewTimeList.Any())
                {

                    int idx = 0;
                    foreach (var rate in Req.RateList)
                    {
                        // R1 = Time 0 - Time 10
                        // R2 = Time 10 - Time 20...以此類推
                        // 但最後兩個不相減
                        var Subtrahend = NewTimeList.Where(o => o.DetailType == rate.DetailType && o.SpecimenID == rate.SpecimenID && o.Time == rate.Subtrahend_Time).FirstOrDefault();
                        var Minuend = NewTimeList.Where(o => o.DetailType == rate.DetailType && o.SpecimenID == rate.SpecimenID && o.Time == rate.Minuend_Time).FirstOrDefault();

                        rate.Subtrahend_TimeUkey = Subtrahend.Ukey;
                        rate.Minuend_TimeUkey = Minuend.Ukey;
                    }
                }
                _Provider.Processe_EvaporationRateTest_Specimen_Rate(Req, UserID);

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

                if (!Req.DetailList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Detail can not be empty.";
                    return model;
                }
                if (!Req.SpecimenList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Specimen List can not be empty.";
                    return model;
                }
                if (!Req.TimeList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Time List can not be empty.";
                    return model;
                }
                if (!Req.RateList.Any())
                {
                    model.Result = false;
                    model.ErrorMessage = "Rate List can not be empty.";
                    return model;
                }

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

                        foreach (var rate in Req.RateList.Where(o => o.DetailType == newSpecimenin.DetailType && o.SpecimenID == newSpecimenin.SpecimenID))
                        {
                            rate.SpecimenUkey = newSpecimenin.Ukey;
                        }

                    }
                }
                _Provider.Processe_EvaporationRateTest_Specimen_Time(Req, UserID, out List<EvaporationRateTest_Specimen_Time> NewTimeList);

                if (NewTimeList.Any())
                {

                    int idx = 0;
                    foreach (var rate in Req.RateList)
                    {
                        // R1 = Time 0 - Time 10
                        // R2 = Time 10 - Time 20...以此類推
                        // 但最後兩個不相減
                        var Subtrahend = NewTimeList.Where(o => o.DetailType == rate.DetailType && o.SpecimenID == rate.SpecimenID && o.Time == rate.Subtrahend_Time).FirstOrDefault();
                        var Minuend = NewTimeList.Where(o => o.DetailType == rate.DetailType && o.SpecimenID == rate.SpecimenID && o.Time == rate.Minuend_Time).FirstOrDefault();

                        rate.Subtrahend_TimeUkey = Subtrahend.Ukey;
                        rate.Minuend_TimeUkey = Minuend.Ukey;
                    }
                }
                _Provider.Processe_EvaporationRateTest_Specimen_Rate(Req, UserID);

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
        public EvaporationRateTest_ViewModel GetReport(string ReportNo, bool isPDF)
        {
            EvaporationRateTest_ViewModel result = new EvaporationRateTest_ViewModel();

            string basefileName = "EvaporationRateTest";
            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

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

                #region Before Sheet

                worksheet.Cells[2, 2] = model.Main.ReportNo;
                worksheet.Cells[2, 7] = model.Main.ReportDateText;

                worksheet.Cells[3, 2] = model.Main.OrderID;
                worksheet.Cells[3, 7] = model.Main.FactoryID;

                worksheet.Cells[4, 2] = model.Main.BrandID;
                worksheet.Cells[4, 7] = model.Main.Article;

                worksheet.Cells[5, 2] = model.Main.StyleID;
                worksheet.Cells[5, 7] = model.Main.FabricColor;

                worksheet.Cells[6, 2] = model.Main.SeasonID;
                worksheet.Cells[6, 7] = model.Main.FabricRefNo;

                worksheet.Cells[7, 1] = $"Before Wash AVERAGE RATE :  {model.Main.BeforeAverageRate} g/h    ";
                worksheet.Cells[7, 2] = $"After Wash AVERAGE RATE :  {model.Main.AfterAverageRate} g/h    ";
                worksheet.Cells[7, 7] = model.Main.Result;

                worksheet.Cells[8, 2] = model.Main.Remark;


                var specimenList = model.SpecimenList.Where(o => o.DetailType == "Before").OrderBy(o => o.SpecimenID).ToList();

                int bonusRow = 0;
                foreach (var specimen in specimenList)
                {
                    // BeforeSpecimen 1 => BeforeSpecimen1
                    string ChartObjectsName = $@"Before{specimen.SpecimenID.Replace(" ", string.Empty)}";
                    var timeListt = model.TimeList.Where(o => o.DetailType == "Before" && o.SpecimenID == specimen.SpecimenID).OrderBy(o => o.Time).ToList();
                    var rateList = model.RateList.Where(o => o.DetailType == "Before" && o.SpecimenID == specimen.SpecimenID).OrderBy(o => o.RateName).ToList();

                    #region Rate 列表處理

                    // 計算每個Rate List的第一Row、最後一Row
                    int startTimeRowBySpecimen = 0;
                    int startRateRowBySpecimen = 0;
                    int lastRateRowBySpecimen = 0;
                    int AverageRateRowBySpecimen = 0;
                    int SpecimenAverageRateRowBySpecimen = 0;

                    switch (specimen.SpecimenID)
                    {
                        case "Specimen 1":
                            startTimeRowBySpecimen = 25;
                            startRateRowBySpecimen = 31;
                            lastRateRowBySpecimen = 33;
                            AverageRateRowBySpecimen = 37;
                            if (rateList.Count() > 3)
                            {
                                bonusRow += rateList.Count() - 3;
                            }
                            break;
                        case "Specimen 2":
                            startTimeRowBySpecimen = 56 + bonusRow;
                            startRateRowBySpecimen = 62 + bonusRow;
                            lastRateRowBySpecimen = 64 + bonusRow;
                            AverageRateRowBySpecimen = 68 + bonusRow;

                            if (rateList.Count() > 3)
                            {
                                bonusRow += rateList.Count() - 3;
                            }
                            break;
                        case "Specimen 3":
                            startTimeRowBySpecimen = 87 + bonusRow;
                            startRateRowBySpecimen = 93 + bonusRow;
                            lastRateRowBySpecimen = 95 + bonusRow;
                            AverageRateRowBySpecimen = 99 + bonusRow;


                            if (rateList.Count() > 3)
                            {
                                bonusRow += rateList.Count() - 3;
                            }

                            // Specimen Average Rate 一個Specimen只會有一個，所以最後再做就好
                            SpecimenAverageRateRowBySpecimen = 103 + bonusRow;
                            break;
                        default:
                            break;
                    }

                    // Rate List 複製Row：若有第4筆，則複製一次；有第5筆，則複製2次
                    for (int i = 0; i < rateList.Count() - 3; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range paste = worksheet.get_Range($"A{lastRateRowBySpecimen + i}", Type.Missing);
                        Microsoft.Office.Interop.Excel.Range copyRow = worksheet.get_Range($"A{lastRateRowBySpecimen}").EntireRow;
                        paste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }

                    // 塞入Rate List的值：前3筆已經有了Rate Name 所以塞值就好，從第4筆開始要額外塞Rate Nmae
                    int addCtn = 0;
                    for (int i = 0; i < rateList.Count(); i++)
                    {
                        string rateName = rateList[i].RateName;
                        int Subtrahend_Time = rateList[i].Subtrahend_Time;
                        int Minuend_Time = rateList[i].Minuend_Time;

                        // R1 ~ R3
                        if (i <= 2)
                        {
                            // Rate Value
                            worksheet.Cells[startRateRowBySpecimen + i, 2] = rateList[i].Value;
                        }
                        else // R4 ~ 
                        {
                            addCtn++;
                            // Rate Name
                            // 範例：R3(Rate of 20~30min)
                            worksheet.Cells[lastRateRowBySpecimen + addCtn, 1] = $"{rateName}(Rate of {Subtrahend_Time}~{Minuend_Time}min)";

                            // Rate Value
                            worksheet.Cells[lastRateRowBySpecimen + addCtn, 2] = rateList[i].Value;
                        }
                    }
                    worksheet.Cells[lastRateRowBySpecimen + addCtn + 4, 2] = specimen.RateAverage;

                    #endregion

                    #region Time 列表處理
                    // Time預設到F欄，超過F欄則動態生成

                    // Time 若有 > 40，則進行複製
                    string lastTimeColumnName = "F";
                    if (timeListt.Count() + 1 > 6) // A欄也要算進去
                    {
                        int baseCtn = 5;
                        int bounsCol = timeListt.Count() - 5;

                        for (int i = 1; i <= bounsCol; i++)
                        {

                            Microsoft.Office.Interop.Excel.Range copyRow = worksheet.get_Range($"{lastTimeColumnName}{startTimeRowBySpecimen}", $"{lastTimeColumnName}{startTimeRowBySpecimen + 2}");
                            string newColName = ConvertNumberToLetter(baseCtn + 1 + i);
                            Microsoft.Office.Interop.Excel.Range paste = worksheet.get_Range($"{newColName}{startTimeRowBySpecimen}", Type.Missing);
                            paste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                        }
                    }

                    // 填值
                    for (int i = 0; i < timeListt.Count(); i++)
                    {
                        int time = timeListt[i].Time;
                        string mass = timeListt[i].Mass.ToString();
                        string evaporation = timeListt[i].Evaporation.ToString();

                        // Time
                        worksheet.Cells[startTimeRowBySpecimen, 2 + i] = time;

                        // Mass
                        worksheet.Cells[startTimeRowBySpecimen + 1, 2 + i] = mass;

                        // Evaporation
                        worksheet.Cells[startTimeRowBySpecimen + 2, 2 + i] = evaporation;
                    }
                    #endregion

                    if (specimen.SpecimenID == "Specimen 3")
                    {
                        // Specimen Rate Average
                        worksheet.Cells[SpecimenAverageRateRowBySpecimen, 2] = model.Main.BeforeAverageRate;
                    }

                    // 折線圖
                    // 取得範本上的圖表集合
                    Microsoft.Office.Interop.Excel.ChartObjects ChartObjects = worksheet.ChartObjects();

                    // 根據Name從集合找出圖表物件
                    Microsoft.Office.Interop.Excel.ChartObject chaartBeforeSpecimen = ChartObjects.Item(ChartObjectsName);
                    Microsoft.Office.Interop.Excel.Chart chart = chaartBeforeSpecimen.Chart;

                    // 找到設定的數列，有幾條線則有幾個數列
                    Microsoft.Office.Interop.Excel.Series series = chaartBeforeSpecimen.Chart.SeriesCollection(1);

                    // 設定X軸Y軸的資料範圍
                    series.XValues = $"='Before wash retest'!$B${startTimeRowBySpecimen}:${lastTimeColumnName}${startTimeRowBySpecimen}";
                    series.Values = $"='Before wash retest'!$B${startTimeRowBySpecimen + 2}:${lastTimeColumnName}${startTimeRowBySpecimen + 2}";

                }

                #endregion
                
                #region After Sheet
                worksheet = excel.ActiveWorkbook.Worksheets[2];

                worksheet.Cells[2, 2] = model.Main.ReportNo;
                worksheet.Cells[2, 7] = model.Main.ReportDateText;

                worksheet.Cells[3, 2] = model.Main.OrderID;
                worksheet.Cells[3, 7] = model.Main.FactoryID;

                worksheet.Cells[4, 2] = model.Main.BrandID;
                worksheet.Cells[4, 7] = model.Main.Article;

                worksheet.Cells[5, 2] = model.Main.StyleID;
                worksheet.Cells[5, 7] = model.Main.FabricColor;

                worksheet.Cells[6, 2] = model.Main.SeasonID;
                worksheet.Cells[6, 7] = model.Main.FabricRefNo;

                worksheet.Cells[7, 1] = $"Before Wash AVERAGE RATE :  {model.Main.BeforeAverageRate} g/h    ";
                worksheet.Cells[7, 2] = $"After Wash AVERAGE RATE :  {model.Main.AfterAverageRate} g/h    ";
                worksheet.Cells[7, 7] = model.Main.Result;

                worksheet.Cells[8, 2] = model.Main.Remark;

                var specimenList2 = model.SpecimenList.Where(o => o.DetailType == "After").OrderBy(o => o.SpecimenID).ToList();

                bonusRow = 0;
                foreach (var specimen in specimenList2)
                {
                    // AfterSpecimen 1 => AfterSpecimen1
                    string ChartObjectsName = $@"After{specimen.SpecimenID.Replace(" ", string.Empty)}";
                    var timeListt = model.TimeList.Where(o => o.DetailType == "After" && o.SpecimenID == specimen.SpecimenID).OrderBy(o => o.Time).ToList();
                    var rateList = model.RateList.Where(o => o.DetailType == "After" && o.SpecimenID == specimen.SpecimenID).OrderBy(o => o.RateName).ToList();


                    #region Rate 列表處理

                    // 計算每個Rate List的第一Row、最後一Row
                    int startTimeRowBySpecimen = 0;
                    int startRateRowBySpecimen = 0;
                    int lastRateRowBySpecimen = 0;
                    int AverageRateRowBySpecimen = 0;
                    int SpecimenAverageRateRowBySpecimen = 0;

                    switch (specimen.SpecimenID)
                    {
                        case "Specimen 1":
                            startTimeRowBySpecimen = 25;
                            startRateRowBySpecimen = 31;
                            lastRateRowBySpecimen = 33;
                            AverageRateRowBySpecimen = 37;
                            if (rateList.Count() > 3)
                            {
                                bonusRow += rateList.Count() - 3;
                            }
                            break;
                        case "Specimen 2":
                            startTimeRowBySpecimen = 56 + bonusRow;
                            startRateRowBySpecimen = 62 + bonusRow;
                            lastRateRowBySpecimen = 64 + bonusRow;
                            AverageRateRowBySpecimen = 68 + bonusRow;

                            if (rateList.Count() > 3)
                            {
                                bonusRow += rateList.Count() - 3;
                            }
                            break;
                        case "Specimen 3":
                            startTimeRowBySpecimen = 87 + bonusRow;
                            startRateRowBySpecimen = 93 + bonusRow;
                            lastRateRowBySpecimen = 95 + bonusRow;
                            AverageRateRowBySpecimen = 99 + bonusRow;

                            // Specimen Average Rate 一個Specimen只會有一個，所以最後再做就好
                            SpecimenAverageRateRowBySpecimen = 103 + bonusRow;

                            if (rateList.Count() > 3)
                            {
                                bonusRow += rateList.Count() - 3;
                            }
                            break;
                        default:
                            break;
                    }

                    // Rate List 複製Row：若有第4筆，則複製一次；有第5筆，則複製2次
                    for (int i = 0; i < rateList.Count() - 3; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range paste = worksheet.get_Range($"A{lastRateRowBySpecimen + i}", Type.Missing);
                        Microsoft.Office.Interop.Excel.Range copyRow = worksheet.get_Range($"A{lastRateRowBySpecimen}").EntireRow;
                        paste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }

                    // 塞入Rate List的值：前3筆已經有了Rate Name 所以塞值就好，從第4筆開始要額外塞Rate Nmae
                    int addCtn = 0;
                    for (int i = 0; i < rateList.Count(); i++)
                    {
                        string rateName = rateList[i].RateName;
                        int Subtrahend_Time = rateList[i].Subtrahend_Time;
                        int Minuend_Time = rateList[i].Minuend_Time;

                        // R1 ~ R3
                        if (i <= 2)
                        {
                            // Rate Value
                            worksheet.Cells[startRateRowBySpecimen + i, 2] = rateList[i].Value;
                        }
                        else // R4 ~ 
                        {
                            addCtn++;
                            // Rate Name
                            // 範例：R3(Rate of 20~30min)
                            worksheet.Cells[lastRateRowBySpecimen + addCtn, 1] = $"{rateName}(Rate of {Subtrahend_Time}~{Minuend_Time}min)";

                            // Rate Value
                            worksheet.Cells[lastRateRowBySpecimen + addCtn, 2] = rateList[i].Value;
                        }
                    }
                    worksheet.Cells[lastRateRowBySpecimen + addCtn + 4, 2] = specimen.RateAverage;

                    #endregion

                    #region Time 列表處理
                    // Time預設到F欄，超過F欄則動態生成

                    // Time 若有 > 40，則進行複製
                    string lastTimeColumnName = "F";
                    if (timeListt.Count() + 1 > 6) // A欄也要算進去
                    {
                        int baseCtn = 5;
                        int bounsCol = timeListt.Count() - 5;

                        for (int i = 1; i <= bounsCol; i++)
                        {

                            Microsoft.Office.Interop.Excel.Range copyRow = worksheet.get_Range($"{lastTimeColumnName}{startTimeRowBySpecimen}", $"{lastTimeColumnName}{startTimeRowBySpecimen + 2}");
                            string newColName = ConvertNumberToLetter(baseCtn + 1 + i);
                            Microsoft.Office.Interop.Excel.Range paste = worksheet.get_Range($"{newColName}{startTimeRowBySpecimen}", Type.Missing);
                            paste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                        }
                    }

                    // 填值
                    for (int i = 0; i < timeListt.Count(); i++)
                    {
                        int time = timeListt[i].Time;
                        string mass = timeListt[i].Mass.ToString();
                        string evaporation = timeListt[i].Evaporation.ToString();

                        // Time
                        worksheet.Cells[startTimeRowBySpecimen, 2 + i] = time;

                        // Mass
                        worksheet.Cells[startTimeRowBySpecimen + 1, 2 + i] = mass;

                        // Evaporation
                        worksheet.Cells[startTimeRowBySpecimen + 2, 2 + i] = evaporation;
                    }
                    #endregion

                    if (specimen.SpecimenID == "Specimen 3")
                    {
                        // Specimen Rate Average
                        worksheet.Cells[SpecimenAverageRateRowBySpecimen, 2] = model.Main.AfterAverageRate;
                    }

                    // 折線圖
                    // 取得範本上的圖表集合
                    Microsoft.Office.Interop.Excel.ChartObjects ChartObjects = worksheet.ChartObjects();

                    // 根據Name從集合找出圖表物件
                    Microsoft.Office.Interop.Excel.ChartObject chaartBeforeSpecimen = ChartObjects.Item(ChartObjectsName);
                    Microsoft.Office.Interop.Excel.Chart chart = chaartBeforeSpecimen.Chart;

                    // 找到設定的數列，有幾條線則有幾個數列
                    Microsoft.Office.Interop.Excel.Series series = chaartBeforeSpecimen.Chart.SeriesCollection(1);

                    // 設定X軸Y軸的資料範圍
                    series.XValues = $"='After wash retest'!$B${startTimeRowBySpecimen}:${lastTimeColumnName}${startTimeRowBySpecimen}";
                    series.Values = $"='After wash retest'!$B${startTimeRowBySpecimen + 2}:${lastTimeColumnName}${startTimeRowBySpecimen + 2}";
                }
                #endregion


                string fileName = $"EvaporationRateTest_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string fullExcelFileName = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                string filePdfName = $"EvaporationRateTest_{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
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

    }

}
