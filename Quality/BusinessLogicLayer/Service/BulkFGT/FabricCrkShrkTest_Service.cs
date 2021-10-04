using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Library;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service
{
    public class FabricCrkShrkTest_Service : IFabricCrkShrkTest_Service
    {
        IFabricCrkShrkTestProvider _FabricCrkShrkTestProvider;
        IScaleProvider _ScaleProvider;
        IOrdersProvider _OrdersProvider;
        IStyleProvider _StyleProvider;
        IFIRLaboratoryProvider _FIRLaboratoryProvider;
        public BaseResult AmendFabricCrkShrkTestCrockingDetail(long ID)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                FabricCrkShrkTestCrocking_Main fabricCrkShrkTestCrocking_Main = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Main(ID);

                if (fabricCrkShrkTestCrocking_Main.CrockingEncdoe == false)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record is not Encode";
                    return baseResult;
                }

                _FabricCrkShrkTestProvider.AmendFabricCrocking(ID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }

            return baseResult;
        }

        public BaseResult AmendFabricCrkShrkTestHeatDetail(long ID)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                FabricCrkShrkTestHeat_Main fabricCrkShrkTestHeat_Main = _FabricCrkShrkTestProvider.GetFabricHeatTest_Main(ID);

                if (fabricCrkShrkTestHeat_Main.HeatEncode == false)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record is not Encode";
                    return baseResult;
                }

                _FabricCrkShrkTestProvider.AmendFabricHeat(ID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }

            return baseResult;
        }

        public BaseResult AmendFabricCrkShrkTestWashDetail(long ID)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                FabricCrkShrkTestWash_Main fabricCrkShrkTestWash_Main = _FabricCrkShrkTestProvider.GetFabricWashTest_Main(ID);

                if (fabricCrkShrkTestWash_Main.WashEncode == false)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record is not Encode";
                    return baseResult;
                }

                _FabricCrkShrkTestProvider.AmendFabricWash(ID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }

            return baseResult;
        }

        public BaseResult EncodeFabricCrkShrkTestCrockingDetail(long ID, string userID, out string testResult)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            testResult = string.Empty;
            try
            {
                FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result = this.GetFabricCrkShrkTestCrocking_Result(ID);

                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingEncdoe == true)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record already Encode";
                    return baseResult;
                }

                if (fabricCrkShrkTestCrocking_Result.Crocking_Detail.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Please test one Roll least.";
                    return baseResult;
                }

                string crockingResult = "Pass";

                if (fabricCrkShrkTestCrocking_Result.Crocking_Detail.Any(s => s.Result.ToUpper() == "FAIL"))
                {
                    crockingResult = "Fail";
                }

                testResult = crockingResult;

                DateTime? crockingDate = fabricCrkShrkTestCrocking_Result.Crocking_Detail.Max(s => s.Inspdate);

                _FabricCrkShrkTestProvider.EncodeFabricCrocking(ID, testResult, crockingDate, userID);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
                return baseResult;
            }
        }

        public BaseResult EncodeFabricCrkShrkTestHeatDetail(long ID, string userID, out string testResult)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            testResult = string.Empty;
            try
            {
                FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result = this.GetFabricCrkShrkTestHeat_Result(ID);

                if (fabricCrkShrkTestHeat_Result.Heat_Main.HeatEncode == true)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record already Encode";
                    return baseResult;
                }

                if (fabricCrkShrkTestHeat_Result.Heat_Detail.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Please test one Roll least.";
                    return baseResult;
                }

                string HeatResult = "Pass";

                if (fabricCrkShrkTestHeat_Result.Heat_Detail.Any(s => s.Result.ToUpper() == "FAIL"))
                {
                    HeatResult = "Fail";
                }

                testResult = HeatResult;

                DateTime? HeatDate = fabricCrkShrkTestHeat_Result.Heat_Detail.Max(s => s.Inspdate);

                _FabricCrkShrkTestProvider.EncodeFabricHeat(ID, testResult, HeatDate, userID);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
                return baseResult;
            }
        }

        public BaseResult EncodeFabricCrkShrkTestWashDetail(long ID, string userID, out string testResult)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            testResult = string.Empty;
            try
            {
                FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result = this.GetFabricCrkShrkTestWash_Result(ID);

                if (fabricCrkShrkTestWash_Result.Wash_Main.WashEncode == true)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record already Encode";
                    return baseResult;
                }

                if (fabricCrkShrkTestWash_Result.Wash_Detail.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Please test one Roll least.";
                    return baseResult;
                }

                string WashResult = "Pass";

                if (fabricCrkShrkTestWash_Result.Wash_Detail.Any(s => s.Result.ToUpper() == "FAIL"))
                {
                    WashResult = "Fail";
                }

                testResult = WashResult;

                DateTime? WashDate = fabricCrkShrkTestWash_Result.Wash_Detail.Max(s => s.Inspdate);

                _FabricCrkShrkTestProvider.EncodeFabricWash(ID, testResult, WashDate, userID);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
                return baseResult;
            }
        }

        public FabricCrkShrkTestCrocking_Result GetFabricCrkShrkTestCrocking_Result(long ID)
        {
            FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result = new FabricCrkShrkTestCrocking_Result();
            try
            {                
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                fabricCrkShrkTestCrocking_Result.Crocking_Main = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Main(ID);

                fabricCrkShrkTestCrocking_Result.Crocking_Detail = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Detail(ID);

                fabricCrkShrkTestCrocking_Result.ID = ID;

                fabricCrkShrkTestCrocking_Result.CrockingTestOption = _FabricCrkShrkTestProvider.GetCrockingTestOption(ID);

                fabricCrkShrkTestCrocking_Result.ScaleIDs = _ScaleProvider.Get().Select(s => s.ID).ToList();

                fabricCrkShrkTestCrocking_Result.Result = true;
            }
            catch (Exception ex)
            {
                fabricCrkShrkTestCrocking_Result.Result = false;
                fabricCrkShrkTestCrocking_Result.ErrorMessage = ex.Message.ToString();
            }

            return fabricCrkShrkTestCrocking_Result;
        }

        public FabricCrkShrkTestHeat_Result GetFabricCrkShrkTestHeat_Result(long ID)
        {
            FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result = new FabricCrkShrkTestHeat_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                fabricCrkShrkTestHeat_Result.Heat_Main = _FabricCrkShrkTestProvider.GetFabricHeatTest_Main(ID);

                fabricCrkShrkTestHeat_Result.Heat_Detail = _FabricCrkShrkTestProvider.GetFabricHeatTest_Detail(ID);

                fabricCrkShrkTestHeat_Result.ID = ID;

                fabricCrkShrkTestHeat_Result.Result = true;
            }
            catch (Exception ex)
            {
                fabricCrkShrkTestHeat_Result.Result = false;

                fabricCrkShrkTestHeat_Result.ErrorMessage = ex.Message.ToString();
            }

            return fabricCrkShrkTestHeat_Result;
        }

        public FabricCrkShrkTestWash_Result GetFabricCrkShrkTestWash_Result(long ID)
        {
            FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result = new FabricCrkShrkTestWash_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                fabricCrkShrkTestWash_Result.Wash_Main = _FabricCrkShrkTestProvider.GetFabricWashTest_Main(ID);

                fabricCrkShrkTestWash_Result.Wash_Detail = _FabricCrkShrkTestProvider.GetFabricWashTest_Detail(ID);

                fabricCrkShrkTestWash_Result.ID = ID;

                fabricCrkShrkTestWash_Result.Result = true;
            }
            catch (Exception ex)
            {
                fabricCrkShrkTestWash_Result.Result = false;
                fabricCrkShrkTestWash_Result.ErrorMessage = ex.Message.ToString();
            }

            return fabricCrkShrkTestWash_Result;
        }

        public FabricCrkShrkTest_Result GetFabricCrkShrkTest_Result(string POID)
        {
            FabricCrkShrkTest_Result result = new FabricCrkShrkTest_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                result = _FabricCrkShrkTestProvider.GetFabricCrkShrkTest_Main(POID);
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public BaseResult SaveFabricCrkShrkTestCrockingDetail(FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                bool isRollDyelotEmpty = fabricCrkShrkTestCrocking_Result.Crocking_Detail.Any(s => MyUtility.Check.Empty(s.Roll) || MyUtility.Check.Empty(s.Dyelot));
                if (isRollDyelotEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Roll and Dyelot cannot be empty.";
                    return baseResult;
                }

                bool isScaleResultEmpty = fabricCrkShrkTestCrocking_Result.Crocking_Detail.Any(s =>
                                                    MyUtility.Check.Empty(s.DryScale) ||
                                                    MyUtility.Check.Empty(s.ResultDry) ||
                                                    MyUtility.Check.Empty(s.WetScale) ||
                                                    MyUtility.Check.Empty(s.ResultWet) ||
                                                    (fabricCrkShrkTestCrocking_Result.CrockingTestOption == 1 &&
                                                        (MyUtility.Check.Empty(s.DryScale_Weft) ||
                                                         MyUtility.Check.Empty(s.ResultDry_Weft) ||
                                                         MyUtility.Check.Empty(s.WetScale_Weft) ||
                                                         MyUtility.Check.Empty(s.ResultWet_Weft)))
                                                    );

                if (isScaleResultEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Scale and Result cannot be empty.";
                    return baseResult;
                }

                bool isInspectorEmpty = fabricCrkShrkTestCrocking_Result.Crocking_Detail.Any(s => MyUtility.Check.Empty(s.Inspector));
                if (isInspectorEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Lab Tech cannot be empty.";
                    return baseResult;
                }

                var listKeyDuplicateItems = fabricCrkShrkTestCrocking_Result
                    .Crocking_Detail.GroupBy(s => new
                    {
                        s.Roll,
                        s.Dyelot,
                    })
                    .Where(groupItem => groupItem.Count() > 1);

                if (listKeyDuplicateItems.Any())
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $@"The following data is duplicated
{listKeyDuplicateItems.Select(s => $"[Roll]{s.Key.Roll}, [Dyelot]{s.Key.Dyelot}").JoinToString(Environment.NewLine)}
";
                    return baseResult;
                }


                //再檢查一次Result
                foreach (FabricCrkShrkTestCrocking_Detail fabricCrkShrkTestCrocking_Detail in fabricCrkShrkTestCrocking_Result.Crocking_Detail)
                {
                    if (fabricCrkShrkTestCrocking_Detail.ResultDry.ToUpper() == "FAIL" ||
                        fabricCrkShrkTestCrocking_Detail.ResultDry_Weft.ToUpper() == "FAIL" ||
                        fabricCrkShrkTestCrocking_Detail.ResultWet.ToUpper() == "FAIL" ||
                        fabricCrkShrkTestCrocking_Detail.ResultWet_Weft.ToUpper() == "FAIL"
                        )
                    {
                        fabricCrkShrkTestCrocking_Detail.Result = "Fail";
                    }
                    else
                    {
                        fabricCrkShrkTestCrocking_Detail.Result = "Pass";
                    }
                }

                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestBeforePicture == null)
                {
                    fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestBeforePicture = new byte[0];
                }

                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestAfterPicture == null)
                {
                    fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestAfterPicture = new byte[0];
                }

                _FabricCrkShrkTestProvider.UpdateFabricCrockingTestDetail(fabricCrkShrkTestCrocking_Result, userID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }

            return baseResult;
        }

        public BaseResult SaveFabricCrkShrkTestHeatDetail(FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                bool isRollDyelotEmpty = fabricCrkShrkTestHeat_Result.Heat_Detail.Any(s => MyUtility.Check.Empty(s.Roll) || MyUtility.Check.Empty(s.Dyelot));
                if (isRollDyelotEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Roll and Dyelot cannot be empty.";
                    return baseResult;
                }

                bool isHorizontalVerticalEmpty = fabricCrkShrkTestHeat_Result.Heat_Detail.Any(s =>
                                                    MyUtility.Check.Empty(s.VerticalOriginal) ||
                                                    MyUtility.Check.Empty(s.HorizontalOriginal) ||
                                                    MyUtility.Check.Empty(s.HorizontalTest1) ||
                                                    MyUtility.Check.Empty(s.HorizontalTest2) ||
                                                    MyUtility.Check.Empty(s.HorizontalTest3) ||
                                                    MyUtility.Check.Empty(s.VerticalTest1) ||
                                                    MyUtility.Check.Empty(s.VerticalTest2) ||
                                                    MyUtility.Check.Empty(s.VerticalTest3)
                                                    );

                if (isHorizontalVerticalEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Horizontal and Vertical cannot be empty.";
                    return baseResult;
                }

                bool isInspectorEmpty = fabricCrkShrkTestHeat_Result.Heat_Detail.Any(s => MyUtility.Check.Empty(s.Inspector));
                if (isInspectorEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Lab Tech cannot be empty.";
                    return baseResult;
                }

                bool isResultEmpty = fabricCrkShrkTestHeat_Result.Heat_Detail.Any(s => MyUtility.Check.Empty(s.Result));
                if (isResultEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Result cannot be empty.";
                    return baseResult;
                }

                var listKeyDuplicateItems = fabricCrkShrkTestHeat_Result
                    .Heat_Detail.GroupBy(s => new
                    {
                        s.Roll,
                        s.Dyelot,
                    })
                    .Where(groupItem => groupItem.Count() > 1);

                if (listKeyDuplicateItems.Any())
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $@"The following data is duplicated
{listKeyDuplicateItems.Select(s => $"[Roll]{s.Key.Roll}, [Dyelot]{s.Key.Dyelot}").JoinToString(Environment.NewLine)}
";
                    return baseResult;
                }

                // 重算HorizontalRate, VerticalRate
                foreach (FabricCrkShrkTestHeat_Detail fabricCrkShrkTestHeat_Detail in fabricCrkShrkTestHeat_Result.Heat_Detail)
                {
                    fabricCrkShrkTestHeat_Detail.HorizontalAverage = (fabricCrkShrkTestHeat_Detail.HorizontalTest1 + fabricCrkShrkTestHeat_Detail.HorizontalTest2 + fabricCrkShrkTestHeat_Detail.HorizontalTest3) / 3;
                    fabricCrkShrkTestHeat_Detail.HorizontalRate = Math.Round((fabricCrkShrkTestHeat_Detail.HorizontalAverage - fabricCrkShrkTestHeat_Detail.HorizontalOriginal) / fabricCrkShrkTestHeat_Detail.HorizontalOriginal * 100, 2);

                    fabricCrkShrkTestHeat_Detail.VerticalAverage = (fabricCrkShrkTestHeat_Detail.VerticalTest1 + fabricCrkShrkTestHeat_Detail.VerticalTest2 + fabricCrkShrkTestHeat_Detail.VerticalTest3) / 3;
                    fabricCrkShrkTestHeat_Detail.VerticalRate = Math.Round((fabricCrkShrkTestHeat_Detail.VerticalAverage - fabricCrkShrkTestHeat_Detail.VerticalOriginal) / fabricCrkShrkTestHeat_Detail.VerticalOriginal * 100, 2);
                }

                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                if (fabricCrkShrkTestHeat_Result.Heat_Main.HeatTestBeforePicture == null)
                {
                    fabricCrkShrkTestHeat_Result.Heat_Main.HeatTestBeforePicture = new byte[0];
                }

                if (fabricCrkShrkTestHeat_Result.Heat_Main.HeatTestAfterPicture == null)
                {
                    fabricCrkShrkTestHeat_Result.Heat_Main.HeatTestAfterPicture = new byte[0];
                }

                _FabricCrkShrkTestProvider.UpdateFabricHeatTestDetail(fabricCrkShrkTestHeat_Result, userID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }

            return baseResult;
        }

        public BaseResult SaveFabricCrkShrkTestMain(FabricCrkShrkTest_Result fabricCrkShrkTest_Result)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                _FabricCrkShrkTestProvider.SaveFabricCrkShrkTest_Main(fabricCrkShrkTest_Result);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }

            return baseResult;
        }

        public BaseResult SaveFabricCrkShrkTestWashDetail(FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                bool isRollDyelotEmpty = fabricCrkShrkTestWash_Result.Wash_Detail.Any(s => MyUtility.Check.Empty(s.Roll) || MyUtility.Check.Empty(s.Dyelot));
                if (isRollDyelotEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Roll and Dyelot cannot be empty.";
                    return baseResult;
                }

                bool isHorizontalVerticalEmpty = fabricCrkShrkTestWash_Result.Wash_Detail.Any(s =>
                                                    MyUtility.Check.Empty(s.VerticalOriginal) ||
                                                    MyUtility.Check.Empty(s.HorizontalOriginal) ||
                                                    MyUtility.Check.Empty(s.HorizontalTest1) ||
                                                    MyUtility.Check.Empty(s.HorizontalTest2) ||
                                                    MyUtility.Check.Empty(s.HorizontalTest3) ||
                                                    MyUtility.Check.Empty(s.VerticalTest1) ||
                                                    MyUtility.Check.Empty(s.VerticalTest2) ||
                                                    MyUtility.Check.Empty(s.VerticalTest3)
                                                    );

                if (isHorizontalVerticalEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Horizontal and Vertical and  cannot be empty.";
                    return baseResult;
                }

                bool isInspectorEmpty = fabricCrkShrkTestWash_Result.Wash_Detail.Any(s => MyUtility.Check.Empty(s.Inspector));
                if (isInspectorEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Lab Tech cannot be empty.";
                    return baseResult;
                }

                bool isResultEmpty = fabricCrkShrkTestWash_Result.Wash_Detail.Any(s => MyUtility.Check.Empty(s.Result));
                if (isResultEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Result cannot be empty.";
                    return baseResult;
                }

                var listKeyDuplicateItems = fabricCrkShrkTestWash_Result
                    .Wash_Detail.GroupBy(s => new
                    {
                        s.Roll,
                        s.Dyelot,
                    })
                    .Where(groupItem => groupItem.Count() > 1);

                if (listKeyDuplicateItems.Any())
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $@"The following data is duplicated
{listKeyDuplicateItems.Select(s => $"[Roll]{s.Key.Roll}, [Dyelot]{s.Key.Dyelot}").JoinToString(Environment.NewLine)}
";
                    return baseResult;
                }

                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                // 重算HorizontalRate, VerticalRate
                foreach (FabricCrkShrkTestWash_Detail fabricCrkShrkTestWash_Detail in fabricCrkShrkTestWash_Result.Wash_Detail)
                {
                    fabricCrkShrkTestWash_Detail.HorizontalAverage = (fabricCrkShrkTestWash_Detail.HorizontalTest1 + fabricCrkShrkTestWash_Detail.HorizontalTest2 + fabricCrkShrkTestWash_Detail.HorizontalTest3) / 3;
                    fabricCrkShrkTestWash_Detail.HorizontalRate = Math.Round((fabricCrkShrkTestWash_Detail.HorizontalAverage - fabricCrkShrkTestWash_Detail.HorizontalOriginal) / fabricCrkShrkTestWash_Detail.HorizontalOriginal * 100, 2);

                    fabricCrkShrkTestWash_Detail.VerticalAverage = (fabricCrkShrkTestWash_Detail.VerticalTest1 + fabricCrkShrkTestWash_Detail.VerticalTest2 + fabricCrkShrkTestWash_Detail.VerticalTest3) / 3;
                    fabricCrkShrkTestWash_Detail.VerticalRate = Math.Round((fabricCrkShrkTestWash_Detail.VerticalAverage - fabricCrkShrkTestWash_Detail.VerticalOriginal) / fabricCrkShrkTestWash_Detail.VerticalOriginal * 100, 2);

                    if (fabricCrkShrkTestWash_Result.Wash_Main.SkewnessOptionID == "1")
                    {
                        fabricCrkShrkTestWash_Detail.HorizontalRate = Math.Abs(fabricCrkShrkTestWash_Detail.HorizontalRate);
                        fabricCrkShrkTestWash_Detail.VerticalRate = Math.Abs(fabricCrkShrkTestWash_Detail.VerticalRate);
                        fabricCrkShrkTestWash_Detail.SkewnessRate = MyUtility.Check.Empty(fabricCrkShrkTestWash_Detail.SkewnessTest1 + fabricCrkShrkTestWash_Detail.SkewnessTest2) ? 0 :
                            (fabricCrkShrkTestWash_Detail.SkewnessTest1 - fabricCrkShrkTestWash_Detail.SkewnessTest2) / (fabricCrkShrkTestWash_Detail.SkewnessTest1 + fabricCrkShrkTestWash_Detail.SkewnessTest2) * 200;
                    }

                    if (fabricCrkShrkTestWash_Result.Wash_Main.SkewnessOptionID == "2")
                    {
                        fabricCrkShrkTestWash_Detail.SkewnessRate = MyUtility.Check.Empty(fabricCrkShrkTestWash_Detail.SkewnessTest3 + fabricCrkShrkTestWash_Detail.SkewnessTest4) ? 0 :
                        (fabricCrkShrkTestWash_Detail.SkewnessTest1 + fabricCrkShrkTestWash_Detail.SkewnessTest2) / (fabricCrkShrkTestWash_Detail.SkewnessTest3 + fabricCrkShrkTestWash_Detail.SkewnessTest4) * 100;
                    }

                    if (fabricCrkShrkTestWash_Result.Wash_Main.SkewnessOptionID == "3")
                    {
                        fabricCrkShrkTestWash_Detail.SkewnessRate = MyUtility.Check.Empty(fabricCrkShrkTestWash_Detail.SkewnessTest2) ? 0 :
                        (fabricCrkShrkTestWash_Detail.SkewnessTest1 / fabricCrkShrkTestWash_Detail.SkewnessTest2) * 100;
                    }
                    
                }

                if (fabricCrkShrkTestWash_Result.Wash_Main.WashTestBeforePicture == null)
                {
                    fabricCrkShrkTestWash_Result.Wash_Main.WashTestBeforePicture = new byte[0];
                }

                if (fabricCrkShrkTestWash_Result.Wash_Main.WashTestAfterPicture == null)
                {
                    fabricCrkShrkTestWash_Result.Wash_Main.WashTestAfterPicture = new byte[0];
                }

                _FabricCrkShrkTestProvider.UpdateFabricWashTestDetail(fabricCrkShrkTestWash_Result, userID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }

            return baseResult;
        }

        public SendMail_Result SendCrockingFailResultMail(string toAddress, string ccAddress, long ID, bool isTest)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricCrkShrkTestProvider.GetCrockingFailMailContentData(ID);
                string mailBody = MailTools.DataTableChangeHtml(dtResult);
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = "Fabric Crocking Test - Test Fail",
                    Body = mailBody
                };
                result = MailTools.SendMail(sendMail_Request, isTest);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.Message.ToString();
            }

            return result;
        }

        public BaseResult ToExcelFabricCrkShrkTestCrockingDetail(long ID, out string excelFileName, bool isTest)
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            BaseResult result = new BaseResult();
            excelFileName = string.Empty;

            try
            {
                string baseFilePath = isTest ? Directory.GetCurrentDirectory() : System.Web.HttpContext.Current.Server.MapPath("~/");
                DataTable dtCrockingDetail = _FabricCrkShrkTestProvider.GetCrockingDetailForReport(ID);
                FabricCrkShrkTestCrocking_Main fabricCrkShrkTestCrocking_Main = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Main(ID);
                string[] columnNames;
                string excelName = string.Empty;

                int crockingTestOption = _FabricCrkShrkTestProvider.GetCrockingTestOption(ID);

                switch (crockingTestOption)
                {
                    case 0:
                        columnNames = new string[] { "Roll", "Dyelot", "DryScale", "WetScale", "Result", "InspDate", "Inspector", "Remark", "LastUpdate" };
                        excelName = baseFilePath + "\\XLT\\FabricCrockingTest.xltx";
                        break;
                    default:
                        columnNames = new string[] { "Roll", "Dyelot", "DryScale", "DryScale_Weft", "WetScale", "WetScale_Weft", "Result", "InspDate", "Inspector", "Remark", "LastUpdate" };
                        excelName = baseFilePath + "\\XLT\\FabricCrockingTestWeftWarp.xltx";
                        break;
                }

                var ret = Array.CreateInstance(typeof(object), dtCrockingDetail.Rows.Count, columnNames.Length) as object[,];
                for (int i = 0; i < dtCrockingDetail.Rows.Count; i++)
                {
                    DataRow row = dtCrockingDetail.Rows[i];
                    for (int j = 0; j < columnNames.Length; j++)
                    {
                        ret[i, j] = row[columnNames[j]];
                    }
                }

                if (dtCrockingDetail.Rows.Count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                // 撈取seasonID
                List<Orders> listOrders = _OrdersProvider.Get(new Orders() { ID = fabricCrkShrkTestCrocking_Main.POID }).ToList();

                string seasonID;

                if (listOrders.Count == 0)
                {
                    seasonID = string.Empty;
                }
                else
                {
                    seasonID = listOrders[0].SeasonID;
                }

                Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(excelName);
                //MyUtility.Excel.CopyToXls(ret, xltFileName: excelName, fileName: string.Empty, openfile: false, headerline: 5, excelAppObj: excel);

                Microsoft.Office.Interop.Excel.Worksheet excelSheets = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表
                excel.Cells[2, 2] = fabricCrkShrkTestCrocking_Main.POID;
                excel.Cells[2, 4] = fabricCrkShrkTestCrocking_Main.SEQ;
                excel.Cells[2, 6] = fabricCrkShrkTestCrocking_Main.ColorID;
                excel.Cells[2, 8] = fabricCrkShrkTestCrocking_Main.StyleID;
                excel.Cells[2, 10] = seasonID;
                excel.Cells[3, 2] = fabricCrkShrkTestCrocking_Main.SCIRefno;
                excel.Cells[3, 4] = fabricCrkShrkTestCrocking_Main.ExportID;
                excel.Cells[3, 6] = fabricCrkShrkTestCrocking_Main.Crocking;
                excel.Cells[3, 8] = fabricCrkShrkTestCrocking_Main.CrockingDate == null ? string.Empty : ((DateTime)fabricCrkShrkTestCrocking_Main.CrockingDate).ToString("yyyy/MM/dd");
                excel.Cells[3, 10] = fabricCrkShrkTestCrocking_Main.BrandID;
                excel.Cells[4, 2] = fabricCrkShrkTestCrocking_Main.Refno;
                excel.Cells[4, 4] = fabricCrkShrkTestCrocking_Main.ArriveQty;
                excel.Cells[4, 6] = fabricCrkShrkTestCrocking_Main.WhseArrival == null ? string.Empty : ((DateTime)fabricCrkShrkTestCrocking_Main.WhseArrival).ToString("yyyy/MM/dd");
                excel.Cells[4, 8] = fabricCrkShrkTestCrocking_Main.Supp;
                excel.Cells[4, 10] = fabricCrkShrkTestCrocking_Main.NonCrocking.ToString();

                int RowIdx = 0;
                foreach (DataRow dr in dtCrockingDetail.Rows)
                {
                    int colIndex = 1;
                    foreach (string col in columnNames)
                    {
                        excel.Cells[RowIdx + 6, colIndex] = dtCrockingDetail.Rows[RowIdx][col].ToString();
                        colIndex++;
                    }

                    RowIdx++;
                }

                excel.Cells.EntireColumn.AutoFit();    // 自動欄寬
                excel.Cells.EntireRow.AutoFit();       ////自動欄高

                #region Save & Show Excel
                excelFileName = $"FabricCrockingTest{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(baseFilePath, "TMP", excelFileName);

                Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                Marshal.ReleaseComObject(excelSheets);
                #endregion

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public BaseResult ToExcelFabricCrkShrkTestHeatDetail(long ID, out string excelFileName, bool isTest)
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            BaseResult result = new BaseResult();
            excelFileName = string.Empty;

            try
            {
                string baseFilePath = isTest ? Directory.GetCurrentDirectory() : System.Web.HttpContext.Current.Server.MapPath("~/");
                DataTable dtHeatDetail = _FabricCrkShrkTestProvider.GetHeatDetailForReport(ID);
                FabricCrkShrkTestHeat_Main fabricCrkShrkTestHeat_Main = _FabricCrkShrkTestProvider.GetFabricHeatTest_Main(ID);

                string excelName = baseFilePath + "\\XLT\\FabricHeatTest.xltx";
                string[] columnNames = new string[]
                {
                "Roll", "Dyelot", "HorizontalOriginal", "VerticalOriginal", "Result", "HorizontalTest1", "HorizontalTest2", "HorizontalTest3", "HorizontalAverage", "HorizontalRate",
                "VerticalTest1", "VerticalTest2", "VerticalTest3", "VerticalAverage", "VerticalRate", "InspDate", "Inspector", "Name", "Remark", "LastUpdate",
                };

                var ret = Array.CreateInstance(typeof(object), dtHeatDetail.Rows.Count, columnNames.Length) as object[,];
                for (int i = 0; i < dtHeatDetail.Rows.Count; i++)
                {
                    DataRow row = dtHeatDetail.Rows[i];
                    for (int j = 0; j < columnNames.Length; j++)
                    {
                        ret[i, j] = row[columnNames[j]];
                    }
                }

                if (dtHeatDetail.Rows.Count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                // 撈seasonID
                // 撈取seasonID
                List<Orders> listOrders = _OrdersProvider.Get(new Orders() { ID = fabricCrkShrkTestHeat_Main.POID }).ToList();

                string seasonID;

                if (listOrders.Count == 0)
                {
                    seasonID = string.Empty;
                }
                else
                {
                    seasonID = listOrders[0].SeasonID;
                }

                Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(excelName);

                Microsoft.Office.Interop.Excel.Worksheet excelSheets = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表
                excelSheets.Cells[2, 2] = fabricCrkShrkTestHeat_Main.POID;
                excelSheets.Cells[2, 4] = fabricCrkShrkTestHeat_Main.SEQ;
                excelSheets.Cells[2, 6] = fabricCrkShrkTestHeat_Main.ColorID;
                excelSheets.Cells[2, 8] = fabricCrkShrkTestHeat_Main.StyleID;
                excelSheets.Cells[2, 10] = seasonID;
                excelSheets.Cells[3, 2] = fabricCrkShrkTestHeat_Main.SCIRefno;
                excelSheets.Cells[3, 4] = fabricCrkShrkTestHeat_Main.ExportID;
                excelSheets.Cells[3, 6] = fabricCrkShrkTestHeat_Main.Heat;
                excelSheets.Cells[3, 8] = fabricCrkShrkTestHeat_Main.HeatDate == null ? string.Empty : ((DateTime)fabricCrkShrkTestHeat_Main.HeatDate).ToString("yyyy/MM/dd");
                excelSheets.Cells[3, 10] = fabricCrkShrkTestHeat_Main.BrandID;
                excelSheets.Cells[4, 2] = fabricCrkShrkTestHeat_Main.Refno;
                excelSheets.Cells[4, 4] = fabricCrkShrkTestHeat_Main.ArriveQty;
                excelSheets.Cells[4, 6] = fabricCrkShrkTestHeat_Main.WhseArrival == null ? string.Empty : ((DateTime)fabricCrkShrkTestHeat_Main.WhseArrival).ToString("yyyy/MM/dd");
                excelSheets.Cells[4, 8] = fabricCrkShrkTestHeat_Main.Supp;
                excelSheets.Cells[4, 10] = fabricCrkShrkTestHeat_Main.NonHeat.ToString();

                int RowIdx = 0;
                foreach (DataRow dr in dtHeatDetail.Rows)
                {
                    int colIndex = 1;
                    foreach (string col in columnNames)
                    {
                        excel.Cells[RowIdx + 6, colIndex] = dtHeatDetail.Rows[RowIdx][col].ToString();
                        colIndex++;
                    }

                    RowIdx++;
                }

                excel.Cells.EntireColumn.AutoFit();    // 自動欄寬
                excel.Cells.EntireRow.AutoFit();       ////自動欄高

                #region Save & Show Excel
                excelFileName = $"FabricHeatTest{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(baseFilePath, "TMP", excelFileName);

                Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                Marshal.ReleaseComObject(excelSheets);
                #endregion
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;

        }

        public BaseResult ToExcelFabricCrkShrkTestWashDetail(long ID, out string excelFileName, bool isTest)
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            BaseResult result = new BaseResult();
            excelFileName = string.Empty;

            try
            {

                string baseFilePath = isTest ? Directory.GetCurrentDirectory() : System.Web.HttpContext.Current.Server.MapPath("~/");

                DataTable dtWashDetail = _FabricCrkShrkTestProvider.GetWashDetailForReport(ID);
                FabricCrkShrkTestWash_Main fabricCrkShrkTestWash_Main = _FabricCrkShrkTestProvider.GetFabricWashTest_Main(ID);

                string excelName = baseFilePath + "\\XLT\\FabricWashTest.xltx";

                string[] columnNames = new string[]
                {
                "Roll", "Dyelot", "HorizontalOriginal", "VerticalOriginal", "Result", "HorizontalTest1", "HorizontalTest2", "HorizontalTest3", "HorizontalAverage", "HorizontalRate",
                "VerticalTest1", "VerticalTest2", "VerticalTest3", "VerticalAverage", "VerticalRate", "SkewnessTest1", "SkewnessTest2", "SkewnessTest3", "SkewnessTest4", "SkewnessRate", "InspDate", "Inspector", "Name", "Remark", "LastUpdate",
                };

                string skewnessOption = fabricCrkShrkTestWash_Main.SkewnessOptionID;

                var ret = Array.CreateInstance(typeof(object), dtWashDetail.Rows.Count, columnNames.Length) as object[,];
                for (int i = 0; i < dtWashDetail.Rows.Count; i++)
                {
                    DataRow row = dtWashDetail.Rows[i];
                    for (int j = 0; j < columnNames.Length; j++)
                    {
                        ret[i, j] = row[columnNames[j]];
                    }
                }

                if (dtWashDetail.Rows.Count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                // 撈seasonID
                List<Orders> listOrders = _OrdersProvider.Get(new Orders() { ID = fabricCrkShrkTestWash_Main.POID }).ToList();

                string seasonID;

                if (listOrders.Count == 0)
                {
                    seasonID = string.Empty;
                }
                else
                {
                    seasonID = listOrders[0].SeasonID;
                }

                Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(excelName);

                Microsoft.Office.Interop.Excel.Worksheet excelSheets = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表

                excelSheets.Cells[2, 2] = fabricCrkShrkTestWash_Main.POID;
                excelSheets.Cells[2, 4] = fabricCrkShrkTestWash_Main.SEQ;
                excelSheets.Cells[2, 6] = fabricCrkShrkTestWash_Main.ColorID;
                excelSheets.Cells[2, 8] = fabricCrkShrkTestWash_Main.StyleID;
                excelSheets.Cells[2, 10] = seasonID;
                excelSheets.Cells[3, 2] = fabricCrkShrkTestWash_Main.SCIRefno;
                excelSheets.Cells[3, 4] = fabricCrkShrkTestWash_Main.ExportID;
                excelSheets.Cells[3, 6] = fabricCrkShrkTestWash_Main.Wash;
                excelSheets.Cells[3, 8] = fabricCrkShrkTestWash_Main.WashDate == null ? string.Empty : ((DateTime)fabricCrkShrkTestWash_Main.WashDate).ToString("yyyy/MM/dd");
                excelSheets.Cells[3, 10] = fabricCrkShrkTestWash_Main.BrandID;
                excelSheets.Cells[4, 2] = fabricCrkShrkTestWash_Main.Refno;
                excelSheets.Cells[4, 4] = fabricCrkShrkTestWash_Main.ArriveQty;
                excelSheets.Cells[4, 6] = fabricCrkShrkTestWash_Main.WhseArrival == null ? string.Empty : ((DateTime)fabricCrkShrkTestWash_Main.WhseArrival).ToString("yyyy/MM/dd");
                excelSheets.Cells[4, 8] = fabricCrkShrkTestWash_Main.Supp;
                excelSheets.Cells[4, 10] = fabricCrkShrkTestWash_Main.NonWash;

                int RowIdx = 0;
                foreach (DataRow dr in dtWashDetail.Rows)
                {
                    int colIndex = 1;
                    foreach (string col in columnNames)
                    {
                        excelSheets.Cells[RowIdx + 6, colIndex] = dtWashDetail.Rows[RowIdx][col].ToString();
                        colIndex++;
                    }

                    RowIdx++;
                }

                // SkewnessOption欄位名稱改變
                switch (skewnessOption)
                {
                    case "1":
                        excelSheets.Cells[5, 16] = "AC";
                        excelSheets.Cells[5, 17] = "BD";
                        break;
                    case "2":
                        excelSheets.Cells[5, 16] = "AA’";
                        excelSheets.Cells[5, 17] = "DD’";
                        excelSheets.Cells[5, 18] = "AB";
                        excelSheets.Cells[5, 19] = "CD";
                        break;
                    case "3":
                        excelSheets.Cells[5, 16] = "AA’";
                        excelSheets.Cells[5, 17] = "AB";
                        break;
                    default:
                        break;
                }

                // 只有2 有4欄，因此1和3藏起來
                if (skewnessOption != "2")
                {
                    excelSheets.get_Range("R:R").EntireColumn.Hidden = true;
                    excelSheets.get_Range("S:S").EntireColumn.Hidden = true;
                }

                excel.Cells.EntireColumn.AutoFit();    // 自動欄寬
                excel.Cells.EntireRow.AutoFit();      ////自動欄高

                // 只有2 有4欄，因此1和3藏起來
                if (skewnessOption != "2")
                {
                    excelSheets.get_Range("R:R").EntireColumn.Hidden = true;
                    excelSheets.get_Range("S:S").EntireColumn.Hidden = true;
                }

                #region Save & Show Excel
                excelFileName = $"FabricWashTest{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(baseFilePath, "TMP", excelFileName);

                Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                Marshal.ReleaseComObject(excelSheets);
                #endregion
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public BaseResult ToPdfFabricCrkShrkTestCrockingDetail(long ID, out string pdfFileName, bool isTest)
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

            BaseResult result = new BaseResult();
            pdfFileName = string.Empty;

            try
            {
                int crockingTestOption = _FabricCrkShrkTestProvider.GetCrockingTestOption(ID);

                if (crockingTestOption == 0)
                {
                    pdfFileName = this.CreateExcelOnlyWetDry(ID, isTest);
                }
                else
                {
                    pdfFileName = this.CreateExcelOnlyWEFTandWARP(ID, isTest);
                }
            }
            catch (Exception ex)
            {

                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        private string CreateExcelOnlyWetDry(long ID, bool isTest)
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _StyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
            _FIRLaboratoryProvider = new FIRLaboratoryProvider(Common.ProductionDataAccessLayer);

            string baseFilePath = isTest ? Directory.GetCurrentDirectory() : System.Web.HttpContext.Current.Server.MapPath("~/");

            FabricCrkShrkTestCrocking_Main fabricCrkShrkTestCrocking_Main = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Main(ID);
            FIR_Laboratory fir_Laboratory = _FIRLaboratoryProvider.Get(new FIR_Laboratory() { ID = ID }).ToList()[0];

            Orders orders = _OrdersProvider.Get(new Orders() { ID = fabricCrkShrkTestCrocking_Main.POID })[0];
            Style style = _StyleProvider.Get(new Style() { Ukey = orders.StyleUkey })[0];

            string submitDate = string.Empty;
            if (!MyUtility.Check.Empty(fir_Laboratory.ReceiveSampleDate))
            {
                submitDate = ((DateTime)fir_Laboratory.ReceiveSampleDate).ToString("yyyy/MM/dd");
            }

            var groupArticle = _FabricCrkShrkTestProvider.GetCrockingArticleForPdfReport(ID).AsEnumerable()
                .GroupBy(s => new
                {
                    Article = s["Article"].ToString(),
                    InspDate = MyUtility.Check.Empty(s["InspDate"]) ? string.Empty : ((DateTime)s["InspDate"]).ToString("yyyy/MM/dd"),
                    Name = s["Name"].ToString()
                });



            Microsoft.Office.Interop.Excel.Application objApp = MyUtility.Excel.ConnectExcel(baseFilePath + "\\XLT\\FabricCrockingTestPDF.xltx");

            objApp.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
            for (int i = 1; i < groupArticle.Count(); i++)
            {
                Microsoft.Office.Interop.Excel.Worksheet worksheet1 = (Microsoft.Office.Interop.Excel.Worksheet)objApp.ActiveWorkbook.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet worksheetn = (Microsoft.Office.Interop.Excel.Worksheet)objApp.ActiveWorkbook.Worksheets[i + 1];

                worksheet1.Copy(worksheetn);
            }

            int j = 0;
            foreach (var groupItem in groupArticle)
            {
                Microsoft.Office.Interop.Excel.Worksheet worksheet = objApp.ActiveWorkbook.Worksheets[j + 1];   // 取得工作表
                worksheet.Cells[4, 3] = submitDate;
                if (!MyUtility.Check.Empty(groupItem.Key.InspDate))
                {
                    worksheet.Cells[4, 5] = groupItem.Key.InspDate;
                }

                worksheet.Cells[6, 9] = groupItem.Key.Article;
                worksheet.Cells[4, 7] = fabricCrkShrkTestCrocking_Main.POID;
                worksheet.Cells[4, 10] = fabricCrkShrkTestCrocking_Main.BrandID;
                worksheet.Cells[6, 3] = fabricCrkShrkTestCrocking_Main.StyleID;
                worksheet.Cells[6, 6] = orders.CustPONO;
                worksheet.Cells[7, 3] = style.StyleName;
                worksheet.Cells[7, 9] = fabricCrkShrkTestCrocking_Main.ArriveQty;
                worksheet.Cells[14, 8] = groupItem.Key.Name;

                for (int i = 1; i < groupItem.Count(); i++)
                {
                    Microsoft.Office.Interop.Excel.Range rngToInsert = worksheet.get_Range("A12:A12", Type.Missing).EntireRow;
                    rngToInsert.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                    Marshal.ReleaseComObject(rngToInsert);
                }

                int k = 0;
                foreach (DataRow detailItemRow in groupItem)
                {
                    worksheet.Cells[11 + k, 2] = fabricCrkShrkTestCrocking_Main.Refno;
                    worksheet.Cells[11 + k, 3] = fabricCrkShrkTestCrocking_Main.ColorID;
                    worksheet.Cells[11 + k, 4] = detailItemRow["Dyelot"];
                    worksheet.Cells[11 + k, 5] = detailItemRow["Roll"];
                    worksheet.Cells[11 + k, 6] = detailItemRow["DryScale"];
                    worksheet.Cells[11 + k, 7] = detailItemRow["ResultDry"];
                    worksheet.Cells[11 + k, 8] = detailItemRow["WetScale"];
                    worksheet.Cells[11 + k, 9] = detailItemRow["ResultWet"];
                    worksheet.Cells[11 + k, 10] = detailItemRow["Remark"];

                    Microsoft.Office.Interop.Excel.Range rg = worksheet.Range[worksheet.Cells[11 + k, 2], worksheet.Cells[11 + k, 10]];

                    // 加框線
                    rg.Borders.LineStyle = 1;
                    rg.Borders.Weight = 3;
                    rg.WrapText = true; // 自動換列
                    rg.Font.Bold = false;

                    // 水平,垂直置中
                    rg.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rg.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    k++;
                }

                // worksheet.get_Range("B9:J9").Font.Bold = true;
                // worksheet.Cells.EntireColumn.AutoFit();
                #region 開始畫格子

                #region 框框數量計算

                int detailCouunt = groupItem.Count();

                // 超過36筆資料，PDF就會跳到下一頁
                // int onePageLimit = 36;

                // 框框數，最多只要顯示4個
                int cubeCount = 0;

                // 每個框框共7個Row高度，因此每多7筆資料，框框就少一個
                if (detailCouunt < 3)
                {
                    cubeCount = 4;
                }

                if (detailCouunt >= 3 && detailCouunt < 10)
                {
                    cubeCount = 3;
                }

                if (detailCouunt >= 10 && detailCouunt < 17)
                {
                    cubeCount = 2;
                }

                if (detailCouunt >= 17 && detailCouunt < 27)
                {
                    cubeCount = 1;
                }

                // 28~44筆資料，未達2頁，但又塞不下一個框框，因此為0

                // 若超過1頁，但未達3頁，還是要畫，第三頁開始不畫

                // 第二頁上面沒有那一堆表格，因此是原本的4 + 表格的10 = 16
                if (detailCouunt >= 43 && detailCouunt < 50)
                {
                    cubeCount = 4;
                }

                if (detailCouunt >= 50 && detailCouunt < 61)
                {
                    cubeCount = 3;
                }

                if (detailCouunt >= 61 && detailCouunt < 72)
                {
                    cubeCount = 2;
                }

                if (detailCouunt >= 72 && detailCouunt < 83)
                {
                    cubeCount = 1;
                }

                // 超過兩頁，會出現第三頁
                if (detailCouunt >= 83)
                {
                    cubeCount = 0;
                }

                #endregion

                // 開始畫
                if (cubeCount > 0)
                {
                    // 作法：先畫第一個，若框框超過一個，就用複製的

                    // 第一個框框上的文字
                    // 16 = 11 + 5
                    worksheet.Cells[16 + groupItem.Count(), 3] = "DRY";
                    worksheet.Cells[16 + groupItem.Count(), 8] = "WET";
                    Microsoft.Office.Interop.Excel.Range rg1 = worksheet.Range[worksheet.Cells[16 + groupItem.Count(), 3], worksheet.Cells[16 + groupItem.Count(), 8]];

                    // 置中
                    rg1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    // 畫框框
                    // 17 = 11 + 6
                    // 24 = 17 + 7
                    rg1 = worksheet.Range[worksheet.Cells[17 + groupItem.Count(), 2], worksheet.Cells[24 + groupItem.Count(), 4]];

                    // 框線設定
                    rg1.BorderAround2(LineStyle: 1);
                    rg1 = worksheet.Range[worksheet.Cells[17 + groupItem.Count(), 7], worksheet.Cells[24 + groupItem.Count(), 9]];
                    rg1.BorderAround2(LineStyle: 1);

                    // 框框旁邊的字
                    rg1 = worksheet.Range[worksheet.Cells[18 + groupItem.Count(), 5], worksheet.Cells[18 + groupItem.Count(), 6]];
                    rg1.Merge(true);
                    worksheet.Cells[18 + groupItem.Count(), 5] = "Ref# : _________________";

                    rg1 = worksheet.Range[worksheet.Cells[19 + groupItem.Count(), 5], worksheet.Cells[19 + groupItem.Count(), 6]];
                    rg1.Merge(true);
                    worksheet.Cells[19 + groupItem.Count(), 5] = "Color  : ________________";

                    rg1 = worksheet.Range[worksheet.Cells[20 + groupItem.Count(), 5], worksheet.Cells[20 + groupItem.Count(), 6]];
                    rg1.Merge(true);
                    worksheet.Cells[20 + groupItem.Count(), 5] = "Roll# : _________________";

                    rg1 = worksheet.Range[worksheet.Cells[21 + groupItem.Count(), 5], worksheet.Cells[21 + groupItem.Count(), 6]];
                    rg1.Merge(true);
                    worksheet.Cells[21 + groupItem.Count(), 5] = "Dyelot# : _______________";

                    rg1 = worksheet.Range[worksheet.Cells[24 + groupItem.Count(), 5], worksheet.Cells[24 + groupItem.Count(), 6]];
                    rg1.Merge(true);
                    worksheet.Cells[24 + groupItem.Count(), 5] = "Grade : ________________";
                    worksheet.Cells[24 + groupItem.Count(), 10] = "_____________";

                    // 選取要被複製的資料
                    rg1 = worksheet.get_Range($"B{17 + groupItem.Count()}:J{24 + groupItem.Count()}").EntireRow;

                    // 根據框框數，資料筆數，決定貼在哪個座標
                    switch (cubeCount)
                    {
                        case 2:
                            Microsoft.Office.Interop.Excel.Range rgX = worksheet.get_Range($"B{17 + groupItem.Count() + 9}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                            rgX.Insert(rg1.Copy(Type.Missing)); // 貼上
                            break;
                        case 3:
                            rgX = worksheet.get_Range($"B{17 + groupItem.Count() + 9}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                            rgX.Insert(rg1.Copy(Type.Missing)); // 貼上
                            Microsoft.Office.Interop.Excel.Range rgY = worksheet.get_Range($"B{17 + groupItem.Count() + 18}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                            rgY.Insert(rg1.Copy(Type.Missing)); // 貼上

                            break;
                        case 4:
                            rgX = worksheet.get_Range($"B{17 + groupItem.Count() + 9}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                            rgX.Insert(rg1.Copy(Type.Missing)); // 貼上
                            rgY = worksheet.get_Range($"B{17 + groupItem.Count() + 18}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                            rgY.Insert(rg1.Copy(Type.Missing)); // 貼上
                            Microsoft.Office.Interop.Excel.Range rgZ = worksheet.get_Range($"B{17 + groupItem.Count() + 27}", Type.Missing).EntireRow; // 選擇要被貼上的位置
                            rgZ.Insert(rg1.Copy(Type.Missing)); // 貼上

                            break;
                        default:
                            break;
                    }
                }

                #endregion

                Marshal.ReleaseComObject(worksheet);
                j++;
            }

            #region Save & Show Excel
            string pdfFileName = $"FabricCrockingTestPDF{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
            string excelFileName = $"FabricCrockingTestPDF{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";

            string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
            string excelPath = Path.Combine(baseFilePath, "TMP", excelFileName);

            objApp.ActiveWorkbook.SaveAs(excelPath);
            objApp.Quit();

            bool isCreatePdfOK = ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
            Marshal.ReleaseComObject(objApp);
            if (!isCreatePdfOK)
            {
                throw new Exception("ConvertToPDF fail");
            }

            return pdfFileName;
            #endregion


        }

        private class PageInfoForPDF
        {
            public int StartRow { get; set; }

            public bool IsSingle { get; set; }
        }

        private List<PageInfoForPDF> GetPageInfo(int firstStartRow, int ttlRowCnt)
        {
            List<PageInfoForPDF> infoForPDFs = new List<PageInfoForPDF>();
            int pagestartRow = firstStartRow;
            bool isSingle = true;

            if (firstStartRow >= 35 && firstStartRow <= 71)
            {
                pagestartRow = 73;
                isSingle = false;
            }

            if (firstStartRow > 71 && ((firstStartRow - 71) % 74) > 37)
            {
                pagestartRow = ((((firstStartRow - 71) / 74) + 1) * 74) + 73;
                isSingle = false;
            }

            infoForPDFs.Add(new PageInfoForPDF { StartRow = pagestartRow, IsSingle = isSingle });

            int removeFirstPageRowCnt = isSingle ? ttlRowCnt - 1 : ttlRowCnt - 2;
            int pageCnt = MyUtility.Convert.GetInt(Math.Ceiling(removeFirstPageRowCnt / 2.0));
            isSingle = false;
            for (int i = 0; i < pageCnt; i++)
            {
                if (i == pageCnt - 1 && (removeFirstPageRowCnt % 2) > 0)
                {
                    isSingle = true;
                }

                if (pagestartRow % 74 != 73)
                {
                    pagestartRow = pagestartRow < 72 ? 73 : ((((pagestartRow - 71) / 74) + 1) * 74) + 73;
                }
                else
                {
                    pagestartRow += 74;
                }

                infoForPDFs.Add(new PageInfoForPDF { StartRow = pagestartRow, IsSingle = isSingle });
            }

            return infoForPDFs;
        }

        private string singleCubeRangeSource = $"A1:N36";
        private string doubleCubeRangeSource = $"A39:N110";

        private string CreateExcelOnlyWEFTandWARP(long ID, bool isTest)
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _StyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
            _FIRLaboratoryProvider = new FIRLaboratoryProvider(Common.ProductionDataAccessLayer);

            string baseFilePath = isTest ? Directory.GetCurrentDirectory() : System.Web.HttpContext.Current.Server.MapPath("~/");

            FabricCrkShrkTestCrocking_Main fabricCrkShrkTestCrocking_Main = _FabricCrkShrkTestProvider.GetFabricCrockingTest_Main(ID);
            FIR_Laboratory fir_Laboratory = _FIRLaboratoryProvider.Get(new FIR_Laboratory() { ID = ID }).ToList()[0];

            Orders orders = _OrdersProvider.Get(new Orders() { ID = fabricCrkShrkTestCrocking_Main.POID })[0];
            Style style = _StyleProvider.Get(new Style() { Ukey = orders.StyleUkey })[0];

            string submitDate = string.Empty;
            if (!MyUtility.Check.Empty(fir_Laboratory.ReceiveSampleDate))
            {
                submitDate = ((DateTime)fir_Laboratory.ReceiveSampleDate).ToString("yyyy/MM/dd");
            }

            var groupArticle = _FabricCrkShrkTestProvider.GetCrockingArticleForPdfReport(ID).AsEnumerable()
                 .GroupBy(s => new
                 {
                     Article = s["Article"].ToString(),
                     InspDate = MyUtility.Check.Empty(s["InspDate"]) ? string.Empty : ((DateTime)s["InspDate"]).ToString("yyyy/MM/dd"),
                     Name = s["Name"].ToString(),
                     Inspector = s["Inspector"].ToString()
                 });

            Microsoft.Office.Interop.Excel.Application objApp = MyUtility.Excel.ConnectExcel(baseFilePath + "\\XLT\\FabricCrockingTestPDFWeftWarp.xltx");
            Microsoft.Office.Interop.Excel.Worksheet worksheetForCopyCube = (Microsoft.Office.Interop.Excel.Worksheet)objApp.ActiveWorkbook.Worksheets[2];
            objApp.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
            for (int i = 1; i < groupArticle.Count(); i++)
            {
                Microsoft.Office.Interop.Excel.Worksheet worksheet1 = (Microsoft.Office.Interop.Excel.Worksheet)objApp.ActiveWorkbook.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet worksheetn = (Microsoft.Office.Interop.Excel.Worksheet)objApp.ActiveWorkbook.Worksheets[i + 1];

                worksheet1.Copy(worksheetn);
            }

            int j = 0;
            foreach (var groupItem in groupArticle)
            {
                Microsoft.Office.Interop.Excel.Worksheet worksheet = objApp.ActiveWorkbook.Worksheets[j + 1];   // 取得工作表
                worksheet.Cells[3, 2] = submitDate;
                if (!MyUtility.Check.Empty(groupItem.Key.InspDate))
                {
                    worksheet.Cells[3, 5] = groupItem.Key.InspDate;
                }

                worksheet.Cells[5, 12] = groupItem.Key.Article;
                worksheet.Cells[3, 9] = fabricCrkShrkTestCrocking_Main.POID;
                worksheet.Cells[3, 13] = fabricCrkShrkTestCrocking_Main.BrandID;
                worksheet.Cells[5, 2] = fabricCrkShrkTestCrocking_Main.StyleID;
                worksheet.Cells[5, 5] = orders.CustPONO;
                worksheet.Cells[6, 2] = style.StyleName;
                worksheet.Cells[6, 12] = fabricCrkShrkTestCrocking_Main.ArriveQty;
                worksheet.Cells[13, 9] = groupItem.Key.Name;

                int k = 10;
                foreach (DataRow detailItemRow in groupItem)
                {
                    worksheet.Cells[k, 1] = fabricCrkShrkTestCrocking_Main.Refno;
                    worksheet.Cells[k, 2] = fabricCrkShrkTestCrocking_Main.ColorID;
                    worksheet.Cells[k, 3] = detailItemRow["Dyelot"];
                    worksheet.Cells[k, 4] = detailItemRow["Roll"];
                    worksheet.Cells[k, 5] = detailItemRow["DryScale"];
                    worksheet.Cells[k, 6] = detailItemRow["ResultDry"];
                    worksheet.Cells[k, 7] = detailItemRow["DryScale_Weft"];
                    worksheet.Cells[k, 8] = detailItemRow["ResultDry_Weft"];
                    worksheet.Cells[k, 9] = detailItemRow["WetScale"];
                    worksheet.Cells[k, 10] = detailItemRow["ResultWet"];
                    worksheet.Cells[k, 11] = detailItemRow["WetScale_Weft"];
                    worksheet.Cells[k, 12] = detailItemRow["ResultWet_Weft"];
                    worksheet.Cells[k, 13] = detailItemRow["Result"];
                    worksheet.Cells[k, 14] = detailItemRow["Remark"];

                    Microsoft.Office.Interop.Excel.Range rg = worksheet.Range[worksheet.Cells[k, 1], worksheet.Cells[k, 14]];

                    // 加框線
                    rg.Borders.LineStyle = 1;
                    rg.Borders.Weight = 3;
                    rg.WrapText = true; // 自動換列
                    rg.Font.Bold = false;

                    // 水平,垂直置中
                    rg.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    rg.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    k++;
                }

                // worksheet.get_Range("B9:J9").Font.Bold = true;
                // worksheet.Cells.EntireColumn.AutoFit();
                #region 開始畫格子

                int firstCubeStartRow = 15 + groupArticle.Count();

                List<PageInfoForPDF> infoForPDFs = this.GetPageInfo(firstCubeStartRow, groupArticle.Count());
                Microsoft.Office.Interop.Excel.Range cubeCopyRange;
                Microsoft.Office.Interop.Excel.Range pastCubeRange;
                foreach (PageInfoForPDF pageInfoForPDF in infoForPDFs)
                {
                    if (pageInfoForPDF.IsSingle)
                    {
                        cubeCopyRange = worksheetForCopyCube.get_Range(this.singleCubeRangeSource, Type.Missing).EntireRow;
                    }
                    else
                    {
                        cubeCopyRange = worksheetForCopyCube.get_Range(this.doubleCubeRangeSource, Type.Missing).EntireRow;
                    }

                    pastCubeRange = worksheet.get_Range($"A{pageInfoForPDF.StartRow}", Type.Missing).EntireRow;
                    pastCubeRange.Insert(cubeCopyRange.Copy(Type.Missing)); // 貼上
                }
                #endregion
                int printPageCountNotIncludeFirst = infoForPDFs[0].IsSingle ? (infoForPDFs.Count - 1) : infoForPDFs.Count;
                int headPageCount = groupArticle.Count() > 58 ? MyUtility.Convert.GetInt(Math.Ceiling((groupItem.Count() - 58) / 74.0)) : 0;
                int lastPageRowNum = lastPageRowNum = 71 + (74 * (printPageCountNotIncludeFirst + headPageCount));

                worksheet.PageSetup.PrintArea = $"$A$1:$N${lastPageRowNum.ToString()}";
                Marshal.ReleaseComObject(worksheet);
                j++;
            }

            worksheetForCopyCube.Delete();

            #region Save & Show Excel
            string pdfFileName = $"FabricCrockingTestPDFWeftWarp{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.pdf";
            string excelFileName = $"FabricCrockingTestPDFWeftWarp{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";

            string pdfPath = Path.Combine(baseFilePath, "TMP", pdfFileName);
            string excelPath = Path.Combine(baseFilePath, "TMP", excelFileName);

            objApp.ActiveWorkbook.SaveAs(excelPath);
            objApp.Quit();

            bool isCreatePdfOK = ConvertToPDF.ExcelToPDF(excelPath, pdfPath);
            Marshal.ReleaseComObject(objApp);
            if (!isCreatePdfOK)
            {
                throw new Exception("ConvertToPDF fail");
            }

            return pdfFileName;
            #endregion
        }

        public SendMail_Result SendHeatFailResultMail(string toAddress, string ccAddress, long ID, bool isTest)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricCrkShrkTestProvider.GetHeatFailMailContentData(ID);
                string mailBody = MailTools.DataTableChangeHtml(dtResult);
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = "Fabric Heat Test - Test Fail",
                    Body = mailBody
                };
                result = MailTools.SendMail(sendMail_Request, isTest);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.Message.ToString();
            }

            return result;
        }

        public SendMail_Result SendWashFailResultMail(string toAddress, string ccAddress, long ID, bool isTest)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricCrkShrkTestProvider.GetWashFailMailContentData(ID);
                string mailBody = MailTools.DataTableChangeHtml(dtResult);
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = "Fabric Wash Test - Test Fail",
                    Body = mailBody
                };
                result = MailTools.SendMail(sendMail_Request, isTest);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.Message.ToString();
            }

            return result;
        }

        public List<string> GetScaleIDs()
        {
            try
            {
                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                return _ScaleProvider.Get().Select(s => s.ID).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
