using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using ClosedXML.Excel;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Web;

namespace BusinessLogicLayer.Service
{
    public class FabricCrkShrkTest_Service : IFabricCrkShrkTest_Service
    {
        IFabricCrkShrkTestProvider _FabricCrkShrkTestProvider;
        IScaleProvider _ScaleProvider;
        IOrdersProvider _OrdersProvider;
        IStyleProvider _StyleProvider;
        IFIRLaboratoryProvider _FIRLaboratoryProvider;
        QualityBrandTestCodeProvider _QualityBrandTestCodeProvider;
        private MailToolsService _MailService;

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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult AmendFabricCrkShrkTestIronDetail(long ID)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                FabricCrkShrkTestIron_Main fabricCrkShrkTestIron_Main = _FabricCrkShrkTestProvider.GetFabricIronTest_Main(ID);

                if (fabricCrkShrkTestIron_Main.IronEncode == false)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record is not Encode";
                    return baseResult;
                }

                _FabricCrkShrkTestProvider.AmendFabricIron(ID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult AmendFabricCrkShrkTestWeightDetail(long ID, string userID)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                FabricCrkShrkTestWeight_Main fabricCrkShrkTestWeight_Main = _FabricCrkShrkTestProvider.GetFabricWeightTest_Main(ID);

                if (fabricCrkShrkTestWeight_Main.WeightEncode == false)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record is not Encode";
                    return baseResult;
                }

                _FabricCrkShrkTestProvider.AmendFabricWeight(ID, userID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
                return baseResult;
            }
        }

        public BaseResult EncodeFabricCrkShrkTestIronDetail(long ID, string userID, out string testResult)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            testResult = string.Empty;
            try
            {
                FabricCrkShrkTestIron_Result fabricCrkShrkTestIron_Result = this.GetFabricCrkShrkTestIron_Result(ID);

                if (fabricCrkShrkTestIron_Result.Iron_Main.IronEncode == true)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record already Encode";
                    return baseResult;
                }

                if (fabricCrkShrkTestIron_Result.Iron_Detail.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Please test one Roll least.";
                    return baseResult;
                }

                string IronResult = "Pass";

                if (fabricCrkShrkTestIron_Result.Iron_Detail.Any(s => s.Result.ToUpper() == "FAIL"))
                {
                    IronResult = "Fail";
                }

                testResult = IronResult;

                DateTime? IronDate = fabricCrkShrkTestIron_Result.Iron_Detail.Max(s => s.Inspdate);

                _FabricCrkShrkTestProvider.EncodeFabricIron(ID, testResult, IronDate, userID);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
                return baseResult;
            }
        }

        public BaseResult EncodeFabricCrkShrkTestWeightDetail(long ID, string userID, out string testResult)
        {
            BaseResult baseResult = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            testResult = string.Empty;
            try
            {
                FabricCrkShrkTestWeight_Result fabricCrkShrkTestWeight_Result = this.GetFabricCrkShrkTestWeight_Result(ID);

                if (fabricCrkShrkTestWeight_Result.Weight_Main.WeightEncode == true)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"This record already Encode";
                    return baseResult;
                }

                if (fabricCrkShrkTestWeight_Result.Weight_Detail.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Please test one Roll least.";
                    return baseResult;
                }

                string WeightResult = "Pass";

                if (fabricCrkShrkTestWeight_Result.Weight_Detail.Any(s => s.Result.ToUpper() == "FAIL"))
                {
                    WeightResult = "Fail";
                }

                testResult = WeightResult;

                DateTime? WeightDate = fabricCrkShrkTestWeight_Result.Weight_Detail.Max(s => s.Inspdate);

                _FabricCrkShrkTestProvider.EncodeFabricWeight(ID, testResult, WeightDate, userID);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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

                DataTable dtResult = _FabricCrkShrkTestProvider.GetCrockingFailMailContentData(ID);
                string Subject = $"Fabric Crocking Test/{fabricCrkShrkTestCrocking_Result.Crocking_Main.POID}/" +
                        $"{dtResult.Rows[0]["Style"]}/" +
                        $"{dtResult.Rows[0]["Refno"]}/" +
                        $"{dtResult.Rows[0]["Color"]}/" +
                        $"{dtResult.Rows[0]["Crocking Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                fabricCrkShrkTestCrocking_Result.Crocking_Main.MailSubject = Subject;

                fabricCrkShrkTestCrocking_Result.Result = true;
            }
            catch (Exception ex)
            {
                fabricCrkShrkTestCrocking_Result.Result = false;
                fabricCrkShrkTestCrocking_Result.ErrorMessage = ex.Message.Replace("'", string.Empty);
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

                DataTable dtResult = _FabricCrkShrkTestProvider.GetHeatFailMailContentData(ID);
                string Subject = $"Fabric Heat Test/{fabricCrkShrkTestHeat_Result.Heat_Main.POID}/" +
                        $"{dtResult.Rows[0]["Style"]}/" +
                        $"{dtResult.Rows[0]["Refno"]}/" +
                        $"{dtResult.Rows[0]["Color"]}/" +
                        $"{dtResult.Rows[0]["Heat Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                fabricCrkShrkTestHeat_Result.Heat_Main.MailSubject = Subject;
                fabricCrkShrkTestHeat_Result.Result = true;
            }
            catch (Exception ex)
            {
                fabricCrkShrkTestHeat_Result.Result = false;

                fabricCrkShrkTestHeat_Result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return fabricCrkShrkTestHeat_Result;
        }
        public FabricCrkShrkTestIron_Result GetFabricCrkShrkTestIron_Result(long ID)
        {
            FabricCrkShrkTestIron_Result fabricCrkShrkTestIron_Result = new FabricCrkShrkTestIron_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                fabricCrkShrkTestIron_Result.Iron_Main = _FabricCrkShrkTestProvider.GetFabricIronTest_Main(ID);

                fabricCrkShrkTestIron_Result.Iron_Detail = _FabricCrkShrkTestProvider.GetFabricIronTest_Detail(ID);

                fabricCrkShrkTestIron_Result.ID = ID;

                DataTable dtResult = _FabricCrkShrkTestProvider.GetIronFailMailContentData(ID);
                string Subject = $"Fabric Iron Test/{fabricCrkShrkTestIron_Result.Iron_Main.POID}/" +
                        $"{dtResult.Rows[0]["Style"]}/" +
                        $"{dtResult.Rows[0]["Refno"]}/" +
                        $"{dtResult.Rows[0]["Color"]}/" +
                        $"{dtResult.Rows[0]["Iron Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                fabricCrkShrkTestIron_Result.Iron_Main.MailSubject = Subject;

                fabricCrkShrkTestIron_Result.Result = true;
            }
            catch (Exception ex)
            {
                fabricCrkShrkTestIron_Result.Result = false;

                fabricCrkShrkTestIron_Result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return fabricCrkShrkTestIron_Result;
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

                DataTable dtResult = _FabricCrkShrkTestProvider.GetWashFailMailContentData(ID);
                string Subject = $"Fabric Wash Test/{fabricCrkShrkTestWash_Result.Wash_Main.POID}/" +
                       $"{dtResult.Rows[0]["Style"]}/" +
                       $"{dtResult.Rows[0]["Refno"]}/" +
                       $"{dtResult.Rows[0]["Color"]}/" +
                       $"{dtResult.Rows[0]["Wash Result"]}/" +
                       $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                fabricCrkShrkTestWash_Result.Wash_Main.MailSubject = Subject;
            }
            catch (Exception ex)
            {
                fabricCrkShrkTestWash_Result.Result = false;
                fabricCrkShrkTestWash_Result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return fabricCrkShrkTestWash_Result;
        }

        public FabricCrkShrkTestWeight_Result GetFabricCrkShrkTestWeight_Result(long ID)
        {
            FabricCrkShrkTestWeight_Result fabricCrkShrkTestWeight_Result = new FabricCrkShrkTestWeight_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                fabricCrkShrkTestWeight_Result.Weight_Main = _FabricCrkShrkTestProvider.GetFabricWeightTest_Main(ID);

                fabricCrkShrkTestWeight_Result.Weight_Detail = _FabricCrkShrkTestProvider.GetFabricWeightTest_Detail(ID);

                fabricCrkShrkTestWeight_Result.ID = ID;

                fabricCrkShrkTestWeight_Result.Result = true;

                fabricCrkShrkTestWeight_Result.Weight_Main.MailSubject = "Fabric Weight Test";
                fabricCrkShrkTestWeight_Result.Weight_Main.MailDesc = "Attachment Is Fabric Weight Test detail data";
            }
            catch (Exception ex)
            {
                fabricCrkShrkTestWeight_Result.Result = false;
                fabricCrkShrkTestWeight_Result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return fabricCrkShrkTestWeight_Result;
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
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
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

                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestPicture1 == null)
                {
                    fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestPicture1 = new byte[0];
                }

                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestPicture2 == null)
                {
                    fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestPicture2 = new byte[0];
                }
                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestPicture3 == null)
                {
                    fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestPicture3 = new byte[0];
                }

                if (fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestPicture4 == null)
                {
                    fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestPicture4 = new byte[0];
                }

                _FabricCrkShrkTestProvider.UpdateFabricCrockingTestDetail(fabricCrkShrkTestCrocking_Result, userID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult SaveFabricCrkShrkTestIronDetail(FabricCrkShrkTestIron_Result fabricCrkShrkTestIron_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                bool isRollDyelotEmpty = fabricCrkShrkTestIron_Result.Iron_Detail.Any(s => MyUtility.Check.Empty(s.Roll) || MyUtility.Check.Empty(s.Dyelot));
                if (isRollDyelotEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Roll and Dyelot cannot be empty.";
                    return baseResult;
                }

                bool isHorizontalVerticalEmpty = fabricCrkShrkTestIron_Result.Iron_Detail.Any(s =>
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

                bool isInspectorEmpty = fabricCrkShrkTestIron_Result.Iron_Detail.Any(s => MyUtility.Check.Empty(s.Inspector));
                if (isInspectorEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Lab Tech cannot be empty.";
                    return baseResult;
                }

                bool isResultEmpty = fabricCrkShrkTestIron_Result.Iron_Detail.Any(s => MyUtility.Check.Empty(s.Result));
                if (isResultEmpty)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Result cannot be empty.";
                    return baseResult;
                }

                var listKeyDuplicateItems = fabricCrkShrkTestIron_Result
                    .Iron_Detail.GroupBy(s => new
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
                foreach (FabricCrkShrkTestIron_Detail fabricCrkShrkTestIron_Detail in fabricCrkShrkTestIron_Result.Iron_Detail)
                {
                    fabricCrkShrkTestIron_Detail.HorizontalAverage = (fabricCrkShrkTestIron_Detail.HorizontalTest1 + fabricCrkShrkTestIron_Detail.HorizontalTest2 + fabricCrkShrkTestIron_Detail.HorizontalTest3) / 3;
                    fabricCrkShrkTestIron_Detail.HorizontalRate = Math.Round((fabricCrkShrkTestIron_Detail.HorizontalAverage - fabricCrkShrkTestIron_Detail.HorizontalOriginal) / fabricCrkShrkTestIron_Detail.HorizontalOriginal * 100, 2);

                    fabricCrkShrkTestIron_Detail.VerticalAverage = (fabricCrkShrkTestIron_Detail.VerticalTest1 + fabricCrkShrkTestIron_Detail.VerticalTest2 + fabricCrkShrkTestIron_Detail.VerticalTest3) / 3;
                    fabricCrkShrkTestIron_Detail.VerticalRate = Math.Round((fabricCrkShrkTestIron_Detail.VerticalAverage - fabricCrkShrkTestIron_Detail.VerticalOriginal) / fabricCrkShrkTestIron_Detail.VerticalOriginal * 100, 2);
                }

                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);

                if (fabricCrkShrkTestIron_Result.Iron_Main.IronTestBeforePicture == null)
                {
                    fabricCrkShrkTestIron_Result.Iron_Main.IronTestBeforePicture = new byte[0];
                }

                if (fabricCrkShrkTestIron_Result.Iron_Main.IronTestAfterPicture == null)
                {
                    fabricCrkShrkTestIron_Result.Iron_Main.IronTestAfterPicture = new byte[0];
                }

                _FabricCrkShrkTestProvider.UpdateFabricIronTestDetail(fabricCrkShrkTestIron_Result, userID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
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
                        fabricCrkShrkTestWash_Detail.HorizontalRate = fabricCrkShrkTestWash_Detail.HorizontalRate;
                        fabricCrkShrkTestWash_Detail.VerticalRate = fabricCrkShrkTestWash_Detail.VerticalRate;
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
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult SaveFabricCrkShrkTestWeightDetail(FabricCrkShrkTestWeight_Result fabricCrkShrkTestWeight_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                if (fabricCrkShrkTestWeight_Result.Weight_Detail.Any(s => MyUtility.Check.Empty(s.Roll) || MyUtility.Check.Empty(s.Dyelot)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "< Roll, Dyelot > can not be empty.";
                    return baseResult;
                }

                if (fabricCrkShrkTestWeight_Result.Weight_Detail.Any(s => s.AverageWeightM2 <= 0))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "< Roll, Dyelot > can not be empty.";
                    return baseResult;
                }

                if (fabricCrkShrkTestWeight_Result.Weight_Detail.Any(s => MyUtility.Check.Empty(s.Inspdate)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "< Inspection Date > can not be empty.";
                    return baseResult;
                }

                if (fabricCrkShrkTestWeight_Result.Weight_Detail.Any(s => MyUtility.Check.Empty(s.Inspector)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "< Lab Tech> can not be empty.";
                    return baseResult;
                }

                var listKeyDuplicateItems = fabricCrkShrkTestWeight_Result
                    .Weight_Detail.GroupBy(s => new
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

                // 重算Difference, Result
                foreach (FabricCrkShrkTestWeight_Detail fabricCrkShrkTestWeight_Detail in fabricCrkShrkTestWeight_Result.Weight_Detail)
                {
                    if (fabricCrkShrkTestWeight_Detail.WeightM2 == 0)
                    {
                        fabricCrkShrkTestWeight_Detail.Result = "Pass";
                        fabricCrkShrkTestWeight_Detail.Difference = 0.0m;
                        continue;
                    }

                    fabricCrkShrkTestWeight_Detail.Difference = MyUtility.Math.Round(((fabricCrkShrkTestWeight_Detail.AverageWeightM2 - fabricCrkShrkTestWeight_Detail.WeightM2) / fabricCrkShrkTestWeight_Detail.WeightM2) * 100, 2);
                    fabricCrkShrkTestWeight_Detail.Result = Math.Abs(fabricCrkShrkTestWeight_Detail.Difference) <= 5 ? "Pass" : "Fail";
                }

                _FabricCrkShrkTestProvider.UpdateFabricWeightTestDetail(fabricCrkShrkTestWeight_Result, userID);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return baseResult;
        }

        public BaseResult ToReport_Crocking(long ID, bool IsPDF, out string excelFileName, string AssignedFineName = "")
        {
            BaseResult result = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);

            List<Crocking_Excel> dataList_Head = _FabricCrkShrkTestProvider.CrockingTest_ToExcel_Head(ID).ToList();
            List<Crocking_Excel> dataList_Body = _FabricCrkShrkTestProvider.CrockingTest_ToExcel_Body(ID).ToList();
            excelFileName = string.Empty;

            if (!dataList_Head.Any())
            {
                result.Result = false;
                result.ErrorMessage = "Data not found!";
                return result;
            }

            // 設定範本檔名
            string basefileName = "FabricCrockingTest";
            switch (dataList_Head.First().BrandID.ToUpper())
            {
                case "ADIDAS":
                case "U.ARMOUR":
                case "NIKE":
                case "GYMSHARK":
                    basefileName = "FabricCrockingTest_ByBrand_A";
                    break;
                case "LLL":
                case "N.FACE":
                case "REI":
                    basefileName = "FabricCrockingTest_ByBrand_B";
                    break;
            }

            // 生成檔案名稱
            var headData = dataList_Head.First();
            string tmpName = $"Fabric Crocking Test_{headData.POID}_{headData.StyleID}_{headData.Refno}_{headData.Color}_{headData.Crocking}_{DateTime.Now:yyyyMMddHHmmss}";
            if(!string.IsNullOrEmpty(AssignedFineName))
            {
                tmpName = AssignedFineName;
            }

            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

            // 開啟範本檔案
            var workbook = new XLWorkbook(openfilepath);
            var worksheet = workbook.Worksheet(1);

            // 填寫表頭資料
            var testCode = _QualityBrandTestCodeProvider.Get(headData.BrandID, "Fabric Crocking & Shrinkage Test-Crocking");
            if (testCode.Any())
            {
                worksheet.Cell("A1").Value = $"Crocking Fastness Test Report({testCode.First().TestCode})";
            }

            worksheet.Cell("C2").Value = headData.ReportNo;
            worksheet.Cell("C3").Value = headData.CrockingReceiveDate?.ToShortDateString();
            worksheet.Cell("G3").Value = headData.SubmitDate;
            worksheet.Cell("C4").Value = headData.SeasonID;
            worksheet.Cell("G4").Value = headData.BrandID;
            worksheet.Cell("C5").Value = headData.StyleID;
            worksheet.Cell("G5").Value = headData.POID;
            worksheet.Cell("C6").Value = headData.Article;
            worksheet.Cell("C7").Value = headData.SCIRefno_Color;
            worksheet.Cell("G8").Value = headData.Color;
            worksheet.Cell("B16").Value = headData.Remark;
            worksheet.Cell("C72").Value = dataList_Head.FirstOrDefault().CrockingInspectorName;
            worksheet.Cell("G72").Value = dataList_Head.FirstOrDefault().CrockingApproverName;

            // 插入圖片
            InsertCrockingImages(worksheet, headData);

            // 動態插入表身資料
            if (dataList_Body.Count > 1)
            {
                int templateRows = 2;
                for (int i = 1; i < dataList_Body.Count; i++)
                {
                    worksheet.Row(13).InsertRowsBelow(templateRows);
                    worksheet.Range("A13:H14").CopyTo(worksheet.Cell(13 + i * templateRows, 1));
                }
            }

            for (int i = 0; i < dataList_Body.Count; i++)
            {
                int rowIdx = i * 2;
                var item = dataList_Body[i];

                worksheet.Cell(13 + rowIdx, 1).Value = item.Roll;
                worksheet.Cell(13 + rowIdx, 2).Value = item.Dyelot;
                worksheet.Cell(13 + rowIdx, 4).Value = item.DryScale;
                worksheet.Cell(13 + rowIdx, 5).Value = item.DryScale_Weft;
                worksheet.Cell(13 + rowIdx, 6).Value = item.WetScale;
                worksheet.Cell(13 + rowIdx, 7).Value = item.WetScale_Weft;

                worksheet.Cell(14 + rowIdx, 4).Value = item.ResultDry;
                worksheet.Cell(14 + rowIdx, 5).Value = item.ResultDry_Weft;
                worksheet.Cell(14 + rowIdx, 6).Value = item.ResultWet;
                worksheet.Cell(14 + rowIdx, 7).Value = item.ResultWet_Weft;
            }

            // Excel 合併 + 塞資料
            #region Title
            string factory = headData.FactoryID.Empty() ? System.Web.HttpContext.Current.Session["FactoryID"].ToString() : headData.FactoryID;

            string FactoryNameEN = _FabricCrkShrkTestProvider.GetFactoryNameEN(factory);
            // 1. 複製第 10 列
            var rowToCopy1 = worksheet.Row(2);

            // 2. 插入一列，將第 8 和第 9 列之間騰出空間
            worksheet.Row(1).InsertRowsAbove(1);

            // 3. 複製格式到新插入的列
            var newRow1 = worksheet.Row(1);
            worksheet.Range("A1:H1").Merge();
            // 設置字體樣式
            var mergedCell = worksheet.Cell("A1");
            mergedCell.Value = FactoryNameEN;
            mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
            mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
            mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            mergedCell.Style.Font.Bold = true;
            // 設置活動儲存格（指標位置）
            worksheet.Cell("A1").SetActive();
            #endregion

            // 去除非法字元
            tmpName = FileNameHelper.SanitizeFileName(tmpName);

            // 儲存報表
            string filexlsx = $"{tmpName}.xlsx";
            string fileNamePDF = $"{tmpName}.pdf";
            string outputPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/TMP/"), filexlsx);
            workbook.SaveAs(outputPath);
            excelFileName = filexlsx;

            //if (IsPDF)
            //{
            //    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
            //    //officeService.ConvertExcelToPdf(outputPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
            //    ConvertToPDF.ExcelToPDF(outputPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/TMP/"), fileNamePDF));
            //    excelFileName = fileNamePDF;
            //}

            result.Result = true;
            return result;
        }
        private void InsertCrockingImages(IXLWorksheet worksheet, Crocking_Excel headData)
        {
            int signatureWidth = 133, signatureHeight = 32;
            int testPictureWidth = 323, testPictureHeight = 255;

            // 根據品牌調整圖片與簽名位置
            switch (headData.BrandID.ToUpper())
            {
                case "ADIDAS":
                case "U.ARMOUR":
                case "NIKE":
                case "GYMSHARK":
                    AddImageToWorksheet(worksheet, headData.InspectorSignature, 74, 3, signatureWidth, signatureHeight);
                    AddImageToWorksheet(worksheet, headData.ApproverSignature, 74, 7, signatureWidth, signatureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture1, 18, 1, testPictureWidth, testPictureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture2, 18, 5, testPictureWidth, testPictureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture3, 46, 1, testPictureWidth, testPictureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture4, 46, 5, testPictureWidth, testPictureHeight);
                    break;

                case "LLL":
                case "N.FACE":
                case "REI":
                    AddImageToWorksheet(worksheet, headData.InspectorSignature, 48, 3, signatureWidth, signatureHeight);
                    AddImageToWorksheet(worksheet, headData.ApproverSignature, 48, 7, signatureWidth, signatureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture1, 18, 1, testPictureWidth, testPictureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture2, 18, 5, testPictureWidth, testPictureHeight);
                    break;

                default: // Default 情況
                    AddImageToWorksheet(worksheet, headData.InspectorSignature, 74, 3, signatureWidth, signatureHeight);
                    AddImageToWorksheet(worksheet, headData.ApproverSignature, 74, 7, signatureWidth, signatureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture1, 18, 1, testPictureWidth, testPictureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture2, 18, 5, testPictureWidth, testPictureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture3, 46, 1, testPictureWidth, testPictureHeight);
                    AddImageToWorksheet(worksheet, headData.CrockingTestPicture4, 46, 5, testPictureWidth, testPictureHeight);
                    break;
            }
        }

        // 儲存報表
        public BaseResult ToReport_Heat(long ID, bool IsPDF, out string excelFileName, string AssignedFineName = "")
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);

            BaseResult result = new BaseResult();
            excelFileName = string.Empty;
            string tmpName = string.Empty;

            try
            {
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                DataTable dtHeatDetail = _FabricCrkShrkTestProvider.GetHeatDetailForReport(ID);
                FabricCrkShrkTestHeat_Main fabricCrkShrkTestHeat_Main = _FabricCrkShrkTestProvider.GetFabricHeatTest_Main(ID);
                var testCode = _QualityBrandTestCodeProvider.Get(fabricCrkShrkTestHeat_Main.BrandID, "Fabric Crocking & Shrinkage Test-Heat");

                string excelName = Path.Combine(baseFilePath, "XLT", "FabricHeatTest.xltx");

                // 如果資料為空，返回錯誤
                if (dtHeatDetail.Rows.Count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                tmpName = $"Fabric Heat Test_{fabricCrkShrkTestHeat_Main.POID}_" +
                          $"{fabricCrkShrkTestHeat_Main.StyleID}_" +
                          $"{fabricCrkShrkTestHeat_Main.Refno}_" +
                          $"{fabricCrkShrkTestHeat_Main.ColorID}_" +
                          $"{fabricCrkShrkTestHeat_Main.Heat}_" +
                          $"{DateTime.Now:yyyyMMddHHmmss}";
                if (AssignedFineName != string.Empty)
                {
                    tmpName = AssignedFineName;
                }

                string seasonID = _OrdersProvider.Get(new Orders() { ID = fabricCrkShrkTestHeat_Main.POID })
                                                 .FirstOrDefault()?.SeasonID ?? string.Empty;

                // 開啟範本檔案
                var workbook = new XLWorkbook(excelName);
                var worksheet = workbook.Worksheet(1);

                // 填寫標題和主表資料
                if (testCode.Any())
                { 
                    worksheet.Cell("A1").Value = $"Quality Heat Test Report({testCode.First().TestCode})";
                }
                worksheet.Cell("B2").Value = fabricCrkShrkTestHeat_Main.ReportNo;

                worksheet.Cell("B3").Value = fabricCrkShrkTestHeat_Main.POID;
                worksheet.Cell("D3").Value = fabricCrkShrkTestHeat_Main.SEQ;
                worksheet.Cell("F3").Value = fabricCrkShrkTestHeat_Main.ColorID;
                worksheet.Cell("H3").Value = fabricCrkShrkTestHeat_Main.StyleID;
                worksheet.Cell("J3").Value = seasonID;

                worksheet.Cell("B4").Value = fabricCrkShrkTestHeat_Main.SCIRefno;
                worksheet.Cell("D4").Value = fabricCrkShrkTestHeat_Main.ExportID;
                worksheet.Cell("F4").Value = fabricCrkShrkTestHeat_Main.Heat;
                worksheet.Cell("H4").Value = fabricCrkShrkTestHeat_Main.HeatDate?.ToString("yyyy/MM/dd");
                worksheet.Cell("J4").Value = fabricCrkShrkTestHeat_Main.BrandID;

                worksheet.Cell("B5").Value = fabricCrkShrkTestHeat_Main.Refno;
                worksheet.Cell("D5").Value = fabricCrkShrkTestHeat_Main.ArriveQty;
                worksheet.Cell("F5").Value = fabricCrkShrkTestHeat_Main.WhseArrival == null ? string.Empty : ((DateTime)fabricCrkShrkTestHeat_Main.WhseArrival).ToString("yyyy/MM/dd");
                worksheet.Cell("H5").Value = fabricCrkShrkTestHeat_Main.Supp;
                worksheet.Cell("J5").Value = fabricCrkShrkTestHeat_Main.NonHeat.ToString();

                worksheet.Cell("B6").Value = fabricCrkShrkTestHeat_Main.HeatReceiveDate.HasValue ? fabricCrkShrkTestHeat_Main.HeatReceiveDate.Value.ToString("yyyy/MM/dd") : string.Empty;

                worksheet.Cell("D30").Value = fabricCrkShrkTestHeat_Main.HeatInspectorName;
                worksheet.Cell("I30").Value = fabricCrkShrkTestHeat_Main.HeatApproverName;

                // 簽名檔圖片
                int signatureWidth = 133, signatureHeight = 32;
                AddImageToWorksheet(worksheet, fabricCrkShrkTestHeat_Main.InspectorSignature, 32, 4, signatureWidth, signatureHeight);
                AddImageToWorksheet(worksheet, fabricCrkShrkTestHeat_Main.ApproverSignature, 32, 9, signatureWidth, signatureHeight);

                // 動態填充明細資料
                int detailStartRow = 8;
                for (int i = 0; i < dtHeatDetail.Rows.Count; i++)
                {
                    // 第一筆資料不用複製
                    if (i > 0)
                    {
                        // 1. 複製上一列的格式
                        var rowToCopy = worksheet.Row(detailStartRow);

                        // 2. 插入新的列
                        worksheet.Row(detailStartRow + i).InsertRowsAbove(1);

                        // 3. 複製格式
                        var newRow = worksheet.Row(detailStartRow + i);

                        rowToCopy.CopyTo(newRow);
                    }

                    var row = dtHeatDetail.Rows[i];
                    for (int col = 0; col < dtHeatDetail.Columns.Count; col++)
                    {
                        switch (row[col].GetType().Name)
                        {
                            case "String":
                                worksheet.Cell(detailStartRow + i, col + 1).Value = row[col].ToString();
                                break;
                            case "Int32":
                                worksheet.Cell(detailStartRow + i, col + 1).Value = row[col].ToValue<Int32>();
                                break;
                            case "Decimal":
                                worksheet.Cell(detailStartRow + i, col + 1).Value = row[col].ToValue<decimal>();
                                break;
                            case "DateTime":
                                worksheet.Cell(detailStartRow + i, col + 1).Value = row[col].ToValue<DateTime>();
                                break;
                            default:
                                break;
                        }
                    }
                }

                // 添加圖片
                if (fabricCrkShrkTestHeat_Main.HeatTestBeforePicture != null)
                {
                    using (var stream = new MemoryStream(fabricCrkShrkTestHeat_Main.HeatTestBeforePicture))
                    {
                        var beforePic = worksheet.AddPicture(stream)
                                                 .MoveTo(worksheet.Cell(10 + dtHeatDetail.Rows.Count, 1))
                                                 .WithSize(323, 255);
                    }
                }

                if (fabricCrkShrkTestHeat_Main.HeatTestAfterPicture != null)
                {
                    using (var stream = new MemoryStream(fabricCrkShrkTestHeat_Main.HeatTestAfterPicture))
                    {
                        var afterPic = worksheet.AddPicture(stream)
                                                .MoveTo(worksheet.Cell(10 + dtHeatDetail.Rows.Count, 7))
                                                .WithSize(430, 340);
                    }
                }

                // Excel 合併 + 塞資料
                #region Title
                string factory = fabricCrkShrkTestHeat_Main.FactoryID.Empty() ? System.Web.HttpContext.Current.Session["FactoryID"].ToString() : fabricCrkShrkTestHeat_Main.FactoryID;

                string FactoryNameEN = _FabricCrkShrkTestProvider.GetFactoryNameEN(factory);
                // 1. 複製第 10 列
                var rowToCopy1 = worksheet.Row(2);

                // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                worksheet.Row(1).InsertRowsAbove(1);

                // 3. 複製格式到新插入的列
                var newRow1 = worksheet.Row(1);
                worksheet.Range("A1:J1").Merge();
                // 設置字體樣式
                var mergedCell = worksheet.Cell("A1");
                mergedCell.Value = FactoryNameEN;
                mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                mergedCell.Style.Font.Bold = true;
                // 設置活動儲存格（指標位置）
                worksheet.Cell("A1").SetActive();
                #endregion

                // 儲存 Excel 檔案
                // 去除非法字元
                tmpName = FileNameHelper.SanitizeFileName(tmpName);
                string filexlsx = $"{tmpName}.xlsx";
                string outputDir = Path.Combine(baseFilePath, "TMP");
                string outputPath = Path.Combine(outputDir, filexlsx);
                workbook.SaveAs(outputPath);
                excelFileName = filexlsx;

                result.Result = true;

                //// 如果需要轉換 PDF
                //if (IsPDF)
                //{
                //    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                //    //officeService.ConvertExcelToPdf(outputPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                //    ConvertToPDF.ExcelToPDF(outputPath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", $"{tmpName}.pdf"));
                //    excelFileName = $"{tmpName}.pdf";
                //}
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }
        public BaseResult ToReport_Iron(long ID, bool IsPDF, out string excelFileName, string AssignedFineName = "")
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            BaseResult result = new BaseResult();
            excelFileName = string.Empty;
            string tmpName = string.Empty;

            try
            {
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");
                DataTable dtIronDetail = _FabricCrkShrkTestProvider.GetIronDetailForReport(ID);
                FabricCrkShrkTestIron_Main fabricCrkShrkTestIron_Main = _FabricCrkShrkTestProvider.GetFabricIronTest_Main(ID);

                 string templatePath = Path.Combine(baseFilePath, "XLT", "FabricIronTest.xltx");

                if (dtIronDetail.Rows.Count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                tmpName = $"Fabric Iron Test_{fabricCrkShrkTestIron_Main.POID}_" +
                          $"{fabricCrkShrkTestIron_Main.StyleID}_" +
                          $"{fabricCrkShrkTestIron_Main.Refno}_" +
                          $"{fabricCrkShrkTestIron_Main.ColorID}_" +
                          $"{fabricCrkShrkTestIron_Main.Iron}_" +
                          $"{DateTime.Now:yyyyMMddHHmmss}";
                if (!string.IsNullOrEmpty(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }
                // 去除非法字元
                tmpName = FileNameHelper.SanitizeFileName(tmpName);

                string seasonID = _OrdersProvider.Get(new Orders { ID = fabricCrkShrkTestIron_Main.POID })
                                    .FirstOrDefault()?.SeasonID ?? string.Empty;

                string outputFilePath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.xlsx");

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填入主表數據
                    worksheet.Cell("B2").Value = fabricCrkShrkTestIron_Main.ReportNo;
                    worksheet.Cell("B3").Value = fabricCrkShrkTestIron_Main.POID;
                    worksheet.Cell("D3").Value = fabricCrkShrkTestIron_Main.SEQ;
                    worksheet.Cell("F3").Value = fabricCrkShrkTestIron_Main.ColorID;
                    worksheet.Cell("H3").Value = fabricCrkShrkTestIron_Main.StyleID;
                    worksheet.Cell("J3").Value = seasonID;
                    worksheet.Cell("B4").Value = fabricCrkShrkTestIron_Main.SCIRefno;
                    worksheet.Cell("D4").Value = fabricCrkShrkTestIron_Main.ExportID;
                    worksheet.Cell("F4").Value = fabricCrkShrkTestIron_Main.Iron;
                    worksheet.Cell("H4").Value = fabricCrkShrkTestIron_Main.IronDate?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell("J4").Value = fabricCrkShrkTestIron_Main.BrandID;
                    worksheet.Cell("B5").Value = fabricCrkShrkTestIron_Main.Refno;
                    worksheet.Cell("D5").Value = fabricCrkShrkTestIron_Main.ArriveQty;
                    worksheet.Cell("F5").Value = fabricCrkShrkTestIron_Main.WhseArrival?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell("H5").Value = fabricCrkShrkTestIron_Main.Supp;
                    worksheet.Cell("J5").Value = fabricCrkShrkTestIron_Main.NonIron.ToString();
                    worksheet.Cell("B6").Value = fabricCrkShrkTestIron_Main.IronReceiveDate?.ToString("yyyy/MM/dd") ?? string.Empty;

                    worksheet.Cell("D30").Value = fabricCrkShrkTestIron_Main.IronInspectorName;
                    worksheet.Cell("I30").Value = fabricCrkShrkTestIron_Main.IronApproverName;

                    // 填入詳細數據
                    int detailStartRow = 8;
                    for (int i = 0; i < dtIronDetail.Rows.Count; i++)
                    {
                        // 第一筆資料不用複製
                        if (i > 0)
                        {
                            // 1. 複製第 10 列
                            var rowToCopy = worksheet.Row(detailStartRow);

                            // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                            worksheet.Row(detailStartRow + i).InsertRowsAbove(1);

                            // 3. 複製格式到新插入的列
                            var newRow = worksheet.Row(detailStartRow + i);

                            rowToCopy.CopyTo(newRow);
                        }

                        var row = dtIronDetail.Rows[i];
                        for (int j = 0; j < dtIronDetail.Columns.Count; j++)
                        {
                            switch (row[j].GetType().Name)
                            {
                                case "String":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToString();
                                    break;
                                case "Int32":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToValue<Int32>();
                                    break;
                                case "Decimal":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToValue<decimal>();
                                    break;
                                case "DateTime":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToValue<DateTime>();
                                    break;
                                default:
                                    break;
                            }
                        }
                    }

                    // 插入圖片（使用共用方法）
                    AddImageToWorksheet(worksheet, fabricCrkShrkTestIron_Main.IronTestBeforePicture, 10 + dtIronDetail.Rows.Count, 1, 323, 255);
                    AddImageToWorksheet(worksheet, fabricCrkShrkTestIron_Main.IronTestAfterPicture, 10 + dtIronDetail.Rows.Count, 7, 323, 255);
                    AddImageToWorksheet(worksheet, fabricCrkShrkTestIron_Main.InspectorSignature, 32 + dtIronDetail.Rows.Count, 4, 100, 24);
                    AddImageToWorksheet(worksheet, fabricCrkShrkTestIron_Main.ApproverSignature, 32 + dtIronDetail.Rows.Count, 9, 100, 24);

                    // Excel 合併 + 塞資料
                    #region Title
                    string factory = fabricCrkShrkTestIron_Main.FactoryID.Empty() ? System.Web.HttpContext.Current.Session["FactoryID"].ToString() : fabricCrkShrkTestIron_Main.FactoryID;

                    string FactoryNameEN = _FabricCrkShrkTestProvider.GetFactoryNameEN(factory);
                    // 1. 複製第 10 列
                    var rowToCopy1 = worksheet.Row(2);

                    // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                    worksheet.Row(1).InsertRowsAbove(1);

                    // 3. 複製格式到新插入的列
                    var newRow1 = worksheet.Row(1);
                    worksheet.Range("A1:J1").Merge();
                    // 設置字體樣式
                    var mergedCell = worksheet.Cell("A1");
                    mergedCell.Value = FactoryNameEN;
                    mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                    mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                    mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    mergedCell.Style.Font.Bold = true;
                    // 設置活動儲存格（指標位置）
                    worksheet.Cell("A1").SetActive();
                    #endregion


                    workbook.SaveAs(outputFilePath);
                    excelFileName = $"{tmpName}.xlsx";
                }

                //if (IsPDF)
                //{
                //    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                //    //officeService.ConvertExcelToPdf(outputFilePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                //    ConvertToPDF.ExcelToPDF(outputFilePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", $"{tmpName}.pdf"));
                //    excelFileName = $"{tmpName}.pdf";
                //}
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }
        public BaseResult ToReport_Wash(long ID, bool IsPDF, out string excelFileName, string AssignedFineName = "")
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
            BaseResult result = new BaseResult();
            excelFileName = string.Empty;
            string tmpName = string.Empty;

            try
            {
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");

                DataTable dtWashDetail = _FabricCrkShrkTestProvider.GetWashDetailForReport(ID);
                FabricCrkShrkTestWash_Main fabricCrkShrkTestWash_Main = _FabricCrkShrkTestProvider.GetFabricWashTest_Main(ID);
                var testCode = _QualityBrandTestCodeProvider.Get(fabricCrkShrkTestWash_Main.BrandID, "Fabric Crocking & Shrinkage Test-Wash");

                string templatePath = Path.Combine(baseFilePath, "XLT", "FabricWashTest.xltx");

                if (dtWashDetail.Rows.Count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                tmpName = $"Fabric Wash Test_{fabricCrkShrkTestWash_Main.POID}_" +
                          $"{fabricCrkShrkTestWash_Main.StyleID}_" +
                          $"{fabricCrkShrkTestWash_Main.Refno}_" +
                          $"{fabricCrkShrkTestWash_Main.ColorID}_" +
                          $"{fabricCrkShrkTestWash_Main.Wash}_" +
                          $"{DateTime.Now:yyyyMMddHHmmss}";
                if (AssignedFineName != "")
                {
                    tmpName = AssignedFineName;
                }
                // 去除非法字元
                tmpName = FileNameHelper.SanitizeFileName(tmpName);

                string seasonID = _OrdersProvider.Get(new Orders { ID = fabricCrkShrkTestWash_Main.POID })
                                    .FirstOrDefault()?.SeasonID ?? string.Empty;

                string outputFilePath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.xlsx");

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 填入主表數據
                    if (testCode.Any())
                    {
                        worksheet.Cell(1, 1).Value = $@"Quality Wash Test Report({testCode.First().TestCode})";
                    }
                    worksheet.Cell("B2").Value = fabricCrkShrkTestWash_Main.ReportNo;
                    worksheet.Cell("B3").Value = fabricCrkShrkTestWash_Main.POID;
                    worksheet.Cell("D3").Value = fabricCrkShrkTestWash_Main.SEQ;
                    worksheet.Cell("F3").Value = fabricCrkShrkTestWash_Main.ColorID;
                    worksheet.Cell("H3").Value = fabricCrkShrkTestWash_Main.StyleID;
                    worksheet.Cell("J3").Value = seasonID;
                    worksheet.Cell("B4").Value = fabricCrkShrkTestWash_Main.SCIRefno;
                    worksheet.Cell("D4").Value = fabricCrkShrkTestWash_Main.ExportID;
                    worksheet.Cell("F4").Value = fabricCrkShrkTestWash_Main.Wash;
                    worksheet.Cell("H4").Value = fabricCrkShrkTestWash_Main.WashDate?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell("J4").Value = fabricCrkShrkTestWash_Main.BrandID;
                    worksheet.Cell("B5").Value = fabricCrkShrkTestWash_Main.Refno;
                    worksheet.Cell("D5").Value = fabricCrkShrkTestWash_Main.ArriveQty;
                    worksheet.Cell("F5").Value = fabricCrkShrkTestWash_Main.WhseArrival?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell("H5").Value = fabricCrkShrkTestWash_Main.Supp;
                    worksheet.Cell("J5").Value = fabricCrkShrkTestWash_Main.NonWash.HasValue ? fabricCrkShrkTestWash_Main.NonWash.Value : false;
                    worksheet.Cell("B6").Value = fabricCrkShrkTestWash_Main.WashReceiveDate?.ToString("yyyy/MM/dd") ?? string.Empty;

                    worksheet.Cell("D30").Value = fabricCrkShrkTestWash_Main.WashInspectorName;
                    worksheet.Cell("I30").Value = fabricCrkShrkTestWash_Main.WashApproverName;

                    // 填入詳細數據
                    int detailStartRow = 8;
                    for (int i = 0; i < dtWashDetail.Rows.Count; i++)
                    {
                        // 第一筆資料不用複製
                        if (i > 0)
                        {
                            // 1. 複製第 10 列
                            var rowToCopy = worksheet.Row(detailStartRow);

                            // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                            worksheet.Row(detailStartRow + i).InsertRowsAbove(1);

                            // 3. 複製格式到新插入的列
                            var newRow = worksheet.Row(detailStartRow + i);

                            rowToCopy.CopyTo(newRow);
                        }

                        var row = dtWashDetail.Rows[i];
                        for (int j = 0; j < dtWashDetail.Columns.Count; j++)
                        {
                            switch (row[j].GetType().Name)
                            {
                                case "String":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToString();
                                    break;
                                case "Int32":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToValue<Int32>();
                                    break;
                                case "Decimal":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToValue<decimal>();
                                    break;
                                case "DateTime":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToValue<DateTime>();
                                    break;
                                default:
                                    break;
                            }
                        }
                    }

                    // SkewnessOption欄位名稱處理
                    switch (fabricCrkShrkTestWash_Main.SkewnessOptionID)
                    {
                        case "1":
                            worksheet.Cell("P6").Value = "AC";
                            worksheet.Cell("Q6").Value = "BD";
                            break;
                        case "2":
                            worksheet.Cell("P6").Value = "AA’";
                            worksheet.Cell("Q6").Value = "DD’";
                            worksheet.Cell("R6").Value = "AB";
                            worksheet.Cell("S6").Value = "CD";
                            break;
                        case "3":
                            worksheet.Cell("P6").Value = "AA’";
                            worksheet.Cell("Q6").Value = "AB";
                            break;
                        default:
                            break;
                    }

                    // 隱藏多餘欄位
                    if (fabricCrkShrkTestWash_Main.SkewnessOptionID != "2")
                    {
                        worksheet.Column("R").Hide();
                        worksheet.Column("S").Hide();
                    }

                    // 插入圖片（使用共用方法）
                    AddImageToWorksheet(worksheet, fabricCrkShrkTestWash_Main.WashTestBeforePicture, 10 + dtWashDetail.Rows.Count, 1, 323, 255);
                    AddImageToWorksheet(worksheet, fabricCrkShrkTestWash_Main.WashTestAfterPicture, 10 + dtWashDetail.Rows.Count, 7, 323, 255);
                    AddImageToWorksheet(worksheet, fabricCrkShrkTestWash_Main.InspectorSignature, 32 + dtWashDetail.Rows.Count, 4, 100, 24);
                    AddImageToWorksheet(worksheet, fabricCrkShrkTestWash_Main.ApproverSignature, 32 + dtWashDetail.Rows.Count, 9, 100, 24);

                    // Excel 合併 + 塞資料
                    #region Title
                    string factory = fabricCrkShrkTestWash_Main.FactoryID.Empty() ? System.Web.HttpContext.Current.Session["FactoryID"].ToString() : fabricCrkShrkTestWash_Main.FactoryID;

                    string FactoryNameEN = _FabricCrkShrkTestProvider.GetFactoryNameEN(factory);
                    // 1. 複製第 10 列
                    var rowToCopy1 = worksheet.Row(2);

                    // 2. 插入一列，將第 8 和第 9 列之間騰出空間
                    worksheet.Row(1).InsertRowsAbove(1);

                    // 3. 複製格式到新插入的列
                    var newRow1 = worksheet.Row(1);
                    worksheet.Range("A1:J1").Merge();
                    // 設置字體樣式
                    var mergedCell = worksheet.Cell("A1");
                    mergedCell.Value = FactoryNameEN;
                    mergedCell.Style.Font.FontName = "Arial";   // 設置字體類型為 Arial
                    mergedCell.Style.Font.FontSize = 25;       // 設置字體大小為 25
                    mergedCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    mergedCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    mergedCell.Style.Font.Bold = true;
                    // 設置活動儲存格（指標位置）
                    worksheet.Cell("A1").SetActive();
                    #endregion
                    workbook.SaveAs(outputFilePath);
                    excelFileName = $"{tmpName}.xlsx";
                }

                //if (IsPDF)
                //{
                //    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                //    //officeService.ConvertExcelToPdf(outputFilePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                //    ConvertToPDF.ExcelToPDF(outputFilePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/TMP/"), $"{tmpName}.pdf"));

                //    excelFileName = $"{tmpName}.pdf";
                //}
                result.Result = true;              
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }
        public BaseResult ToReport_Weight(long ID, bool IsPDF, out string excelFileName, string AssignedFineName = "")
        {
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _OrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
            BaseResult result = new BaseResult();
            excelFileName = string.Empty;
            string tmpName = string.Empty;

            try
            {
                string baseFilePath = System.Web.HttpContext.Current.Server.MapPath("~/");

                DataTable dtWeightDetail = _FabricCrkShrkTestProvider.GetWeightDetailForReport(ID);
                FabricCrkShrkTestWeight_Main fabricCrkShrkTestWeight_Main = _FabricCrkShrkTestProvider.GetFabricWeightTest_Main(ID);

                string templatePath = Path.Combine(baseFilePath, "XLT", "FabricWeightTest.xltx");

                if (dtWeightDetail.Rows.Count == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }
                Guid newGuid = Guid.NewGuid();

                tmpName = "FabricWeightTest_" +
                          $"{DateTime.Now:yyyyMMddHHmmss}_" +
                          $"{newGuid.ToString()}";
                if (AssignedFineName != "")
                {
                    tmpName = AssignedFineName;
                }
                // 去除非法字元
                tmpName = FileNameHelper.SanitizeFileName(tmpName);

                string seasonID = _OrdersProvider.Get(new Orders { ID = fabricCrkShrkTestWeight_Main.POID })
                                    .FirstOrDefault()?.SeasonID ?? string.Empty;

                string outputFilePath = Path.Combine(baseFilePath, "TMP", $"{tmpName}.xlsx");

                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    worksheet.Cell("B2").Value = fabricCrkShrkTestWeight_Main.POID;
                    worksheet.Cell("D2").Value = fabricCrkShrkTestWeight_Main.SEQ;
                    worksheet.Cell("F2").Value = fabricCrkShrkTestWeight_Main.ColorID;
                    worksheet.Cell("H2").Value = fabricCrkShrkTestWeight_Main.StyleID;
                    worksheet.Cell("J2").Value = seasonID;
                    worksheet.Cell("B3").Value = fabricCrkShrkTestWeight_Main.SCIRefno;
                    worksheet.Cell("D3").Value = fabricCrkShrkTestWeight_Main.WeightEncode.ToString();
                    worksheet.Cell("F3").Value = fabricCrkShrkTestWeight_Main.Weight;
                    worksheet.Cell("H3").Value = fabricCrkShrkTestWeight_Main.WeightDate?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell("J3").Value = fabricCrkShrkTestWeight_Main.BrandID;
                    worksheet.Cell("B4").Value = fabricCrkShrkTestWeight_Main.Refno;
                    worksheet.Cell("D4").Value = fabricCrkShrkTestWeight_Main.ArriveQty;
                    worksheet.Cell("F4").Value = fabricCrkShrkTestWeight_Main.WhseArrival?.ToString("yyyy/MM/dd") ?? string.Empty;
                    worksheet.Cell("H4").Value = fabricCrkShrkTestWeight_Main.Supp;
                    worksheet.Cell("J4").Value = fabricCrkShrkTestWeight_Main.ExportID;

                    // 填入詳細數據
                    int detailStartRow = 6;
                    for (int i = 0; i < dtWeightDetail.Rows.Count; i++)
                    {
                        // 第一筆資料不用複製
                        if (i > 0)
                        {
                            // 1. 複製第 10 列
                            var rowToCopy = worksheet.Row(detailStartRow);

                            // 2. 插入一列，將第 5 和第 6 列之間騰出空間
                            worksheet.Row(detailStartRow + i).InsertRowsAbove(1);

                            // 3. 複製格式到新插入的列
                            var newRow = worksheet.Row(detailStartRow + i);

                            rowToCopy.CopyTo(newRow);
                        }

                        var row = dtWeightDetail.Rows[i];
                        for (int j = 0; j < dtWeightDetail.Columns.Count; j++)
                        {
                            switch (row[j].GetType().Name)
                            {
                                case "String":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToString();
                                    break;
                                case "Int32":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToValue<Int32>();
                                    break;
                                case "Decimal":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToValue<decimal>();
                                    break;
                                case "DateTime":
                                    worksheet.Cell(detailStartRow + i, j + 1).Value = row[j].ToValue<DateTime>();
                                    break;
                                default:
                                    break;
                            }
                        }
                    }

                    workbook.SaveAs(outputFilePath);
                    excelFileName = $"{tmpName}.xlsx";
                }

                if (IsPDF)
                {
                    //LibreOfficeService officeService = new LibreOfficeService(@"C:\Program Files\LibreOffice\program\");
                    //officeService.ConvertExcelToPdf(outputFilePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP"));
                    ConvertToPDF.ExcelToPDF(outputFilePath, Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/TMP/"), $"{tmpName}.pdf"));

                    excelFileName = $"{tmpName}.pdf";
                }
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return result;
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

            if (firstStartRow >= 35 && firstStartRow <= 75)
            {
                pagestartRow = 77;
                isSingle = false;
            }

            if (firstStartRow > 75 && ((firstStartRow - 75) % 82) > 37)
            {
                pagestartRow = ((((firstStartRow - 75) / 82) + 1) * 82) + 81;
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

                if (pagestartRow % 82 != 81)
                {
                    pagestartRow = pagestartRow < 72 ? 81 : ((((pagestartRow - 75) / 82) + 1) * 82) + 81;
                }
                else
                {
                    pagestartRow += 82;
                }

                infoForPDFs.Add(new PageInfoForPDF { StartRow = pagestartRow, IsSingle = isSingle });
            }

            return infoForPDFs;
        }

        private string singleCubeRangeSource = $"A1:N36";
        private string doubleCubeRangeSource = $"A39:N110";

        public SendMail_Result SendHeatFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricCrkShrkTestProvider.GetHeatFailMailContentData(ID);
                string name = $"Fabric Heat Test_{OrderID}_" +
                        $"{dtResult.Rows[0]["Style"]}_" +
                        $"{dtResult.Rows[0]["Refno"]}_" +
                        $"{dtResult.Rows[0]["Color"]}_" +
                        $"{dtResult.Rows[0]["Heat Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                BaseResult baseResult = ToReport_Heat(ID, true, out string excelFileName, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", excelFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = $"Fabric Heat Test/{OrderID}/" +
                        $"{dtResult.Rows[0]["Style"]}/" +
                        $"{dtResult.Rows[0]["Refno"]}/" +
                        $"{dtResult.Rows[0]["Color"]}/" +
                        $"{dtResult.Rows[0]["Heat Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Fabric Crocking & Shrinkage Test",
                    OrderID = OrderID,
                };

                if (!string.IsNullOrEmpty(Subject))
                {
                    sendMail_Request.Subject = Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dtResult, comment, buyReadyDate, Body, out AlternateView plainView);

                sendMail_Request.Body = mailBody;
                sendMail_Request.alternateView = plainView;

                result = MailTools.SendMail(sendMail_Request);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }
        public SendMail_Result SendIronFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricCrkShrkTestProvider.GetIronFailMailContentData(ID);
                string name = $"Fabric Iron Test_{OrderID}_" +
                        $"{dtResult.Rows[0]["Style"]}_" +
                        $"{dtResult.Rows[0]["Refno"]}_" +
                        $"{dtResult.Rows[0]["Color"]}_" +
                        $"{dtResult.Rows[0]["Iron Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                BaseResult baseResult = ToReport_Iron(ID, true, out string excelFileName, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", excelFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = $"Fabric Iron Test/{OrderID}/" +
                        $"{dtResult.Rows[0]["Style"]}/" +
                        $"{dtResult.Rows[0]["Refno"]}/" +
                        $"{dtResult.Rows[0]["Color"]}/" +
                        $"{dtResult.Rows[0]["Iron Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Fabric Crocking & Shrinkage Test",
                    OrderID = OrderID,
                };

                if (!string.IsNullOrEmpty(Subject))
                {
                    sendMail_Request.Subject = Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dtResult, comment, buyReadyDate, Body, out AlternateView plainView);

                sendMail_Request.Body = mailBody;
                sendMail_Request.alternateView = plainView;

                result = MailTools.SendMail(sendMail_Request);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }
        public SendMail_Result SendWashFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricCrkShrkTestProvider.GetWashFailMailContentData(ID);
                string name = $"Fabric Wash Test_{OrderID}_" +
                        $"{dtResult.Rows[0]["Style"]}_" +
                        $"{dtResult.Rows[0]["Refno"]}_" +
                        $"{dtResult.Rows[0]["Color"]}_" +
                        $"{dtResult.Rows[0]["Wash Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                BaseResult baseResult = ToReport_Wash(ID, true, out string excelFileName, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", excelFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = $"Fabric Wash Test/{OrderID}/" +
                        $"{dtResult.Rows[0]["Style"]}/" +
                        $"{dtResult.Rows[0]["Refno"]}/" +
                        $"{dtResult.Rows[0]["Color"]}/" +
                        $"{dtResult.Rows[0]["Wash Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Fabric Crocking & Shrinkage Test",
                    OrderID = OrderID,
                };

                if (!string.IsNullOrEmpty(Subject))
                {
                    sendMail_Request.Subject = Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dtResult, comment, buyReadyDate, Body, out AlternateView plainView);

                sendMail_Request.Body = mailBody;
                sendMail_Request.alternateView = plainView;

                result = MailTools.SendMail(sendMail_Request);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }
        public SendMail_Result SendWeightFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                string name = $"Fabric Weight Test";

                BaseResult baseResult = ToReport_Weight(ID, true, out string excelFileName, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", excelFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = "Fabric Weight Test",
                    Body = "Attachment Is Fabric Weight Test detail data",
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Fabric Crocking & Shrinkage Test",
                    OrderID = OrderID,
                };

                if (!string.IsNullOrEmpty(Subject))
                {
                    sendMail_Request.Subject = Subject;
                }

                if (!string.IsNullOrEmpty(Body))
                {
                    sendMail_Request.Body = Body;
                }

                _MailService = new MailToolsService();

                result = MailTools.SendMail(sendMail_Request);

            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }

        public SendMail_Result SendCrockingFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files)
        {
            SendMail_Result result = new SendMail_Result();
            try
            {
                _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricCrkShrkTestProvider.GetCrockingFailMailContentData(ID);

                string name = $"Fabric Crocking Test_{OrderID}_" +
                        $"{dtResult.Rows[0]["Style"]}_" +
                        $"{dtResult.Rows[0]["Refno"]}_" +
                        $"{dtResult.Rows[0]["Color"]}_" +
                        $"{dtResult.Rows[0]["Crocking Result"]}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                BaseResult baseResult = ToReport_Crocking(ID, true, out string excelFileName, name);
                string FileName = baseResult.Result ? Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", excelFileName) : string.Empty;
                SendMail_Request sendMail_Request = new SendMail_Request()
                {
                    To = toAddress,
                    CC = ccAddress,
                    Subject = $"Fabric Crocking Test/{OrderID}/" +
                        $"{dtResult.Rows[0]["Style"]}/" +
                        $"{dtResult.Rows[0]["Refno"]}/" +
                        $"{dtResult.Rows[0]["Color"]}/" +
                        $"{dtResult.Rows[0]["Crocking Result"]}/" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}",
                    //Body = mailBody,
                    //alternateView = plainView,
                    FileonServer = new List<string> { FileName },
                    FileUploader = Files,
                    IsShowAIComment = true,
                    AICommentType = "Fabric Crocking & Shrinkage Test",
                    OrderID = OrderID,
                };

                if (!string.IsNullOrEmpty(Subject))
                {
                    sendMail_Request.Subject = Subject;
                }

                _MailService = new MailToolsService();
                string comment = _MailService.GetAICommet(sendMail_Request);
                string buyReadyDate = _MailService.GetBuyReadyDate(sendMail_Request);
                string mailBody = MailTools.DataTableChangeHtml(dtResult, comment, buyReadyDate, Body, out AlternateView plainView);

                sendMail_Request.Body = mailBody;
                sendMail_Request.alternateView = plainView;


                result = MailTools.SendMail(sendMail_Request);
            }
            catch (Exception ex)
            {
                result.result = false;
                result.resultMsg = ex.Message.Replace("'", string.Empty);
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
