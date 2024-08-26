using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using Library;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using Org.BouncyCastle.Ocsp;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;
using Style = DatabaseObject.ProductionDB.Style;

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

                tmpName = $"Fabric Heat Test_{fabricCrkShrkTestHeat_Main.POID}_" +
                        $"{fabricCrkShrkTestHeat_Main.StyleID}_" +
                        $"{fabricCrkShrkTestHeat_Main.Refno}_" +
                        $"{fabricCrkShrkTestHeat_Main.ColorID}_" +
                        $"{fabricCrkShrkTestHeat_Main.Heat}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";


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
                if (testCode.Any())
                {
                    excelSheets.Cells[1, 1] = $@"Quality Heat Test Report({testCode.FirstOrDefault().TestCode})";
                }
                excelSheets.Cells[2, 2] = fabricCrkShrkTestHeat_Main.ReportNo;
                excelSheets.Cells[3, 2] = fabricCrkShrkTestHeat_Main.POID;
                excelSheets.Cells[3, 4] = fabricCrkShrkTestHeat_Main.SEQ;
                excelSheets.Cells[3, 6] = fabricCrkShrkTestHeat_Main.ColorID;
                excelSheets.Cells[3, 8] = fabricCrkShrkTestHeat_Main.StyleID;
                excelSheets.Cells[3, 10] = seasonID;
                excelSheets.Cells[4, 2] = fabricCrkShrkTestHeat_Main.SCIRefno;
                excelSheets.Cells[4, 4] = fabricCrkShrkTestHeat_Main.ExportID;
                excelSheets.Cells[4, 6] = fabricCrkShrkTestHeat_Main.Heat;
                excelSheets.Cells[4, 8] = fabricCrkShrkTestHeat_Main.HeatDate == null ? string.Empty : ((DateTime)fabricCrkShrkTestHeat_Main.HeatDate).ToString("yyyy/MM/dd");
                excelSheets.Cells[4, 10] = fabricCrkShrkTestHeat_Main.BrandID;
                excelSheets.Cells[5, 2] = fabricCrkShrkTestHeat_Main.Refno;
                excelSheets.Cells[5, 4] = fabricCrkShrkTestHeat_Main.ArriveQty;
                excelSheets.Cells[5, 6] = fabricCrkShrkTestHeat_Main.WhseArrival == null ? string.Empty : ((DateTime)fabricCrkShrkTestHeat_Main.WhseArrival).ToString("yyyy/MM/dd");
                excelSheets.Cells[5, 8] = fabricCrkShrkTestHeat_Main.Supp;
                excelSheets.Cells[5, 10] = fabricCrkShrkTestHeat_Main.NonHeat.ToString();

                int defaultRowCount = 10;
                int otherCount = dtHeatDetail.Rows.Count - defaultRowCount;
                if (otherCount > 0)
                {
                    // Rate List 複製Row：若有第4筆，則複製一次；有第5筆，則複製2次
                    for (int i = 0; i < otherCount; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range paste = excelSheets.get_Range($"A{defaultRowCount + i}", Type.Missing);
                        Microsoft.Office.Interop.Excel.Range copyRow = excelSheets.get_Range($"A{defaultRowCount}").EntireRow;
                        paste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }
                }

                int RowIdx = 0;
                foreach (DataRow dr in dtHeatDetail.Rows)
                {
                    int colIndex = 1;
                    foreach (string col in columnNames)
                    {
                        excel.Cells[RowIdx + 7, colIndex] = dtHeatDetail.Rows[RowIdx][col].ToString();
                        colIndex++;
                    }

                    RowIdx++;
                }

                #region 添加圖片
                Excel.Range cellBeforePicture = excelSheets.Cells[22 + otherCount, 1];
                if (fabricCrkShrkTestHeat_Main.HeatTestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(fabricCrkShrkTestHeat_Main.HeatTestBeforePicture, fabricCrkShrkTestHeat_Main.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    excelSheets.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 323, 255);
                }

                Excel.Range cellAfterPicture = excelSheets.Cells[22 + otherCount, 7];
                if (fabricCrkShrkTestHeat_Main.HeatTestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(fabricCrkShrkTestHeat_Main.HeatTestAfterPicture, fabricCrkShrkTestHeat_Main.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    excelSheets.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 2, 323, 255);
                }
                #endregion

                excel.Cells.EntireColumn.AutoFit();    // 自動欄寬
                excel.Cells.EntireRow.AutoFit();       ////自動欄高

                #region Save & Show Excel

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }
                char[] invalidChars = Path.GetInvalidFileNameChars();

                foreach (char invalidChar in invalidChars)
                {
                    tmpName = tmpName.Replace(invalidChar.ToString(), "");
                }

                string filexlsx = tmpName + ".xlsx";
                string fileNamePDF = tmpName + ".pdf";

                string filepathpdf = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF);
                string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);


                Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(filepath);

                excelFileName = filexlsx;

                if (IsPDF)
                {
                    if (ConvertToPDF.ExcelToPDF(filepath, filepathpdf))
                    {
                        excelFileName = fileNamePDF;
                        result.Result = true;
                    }
                    else
                    {
                        result.ErrorMessage = "Convert To PDF Fail";
                        result.Result = false;
                    }

                }

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                Marshal.ReleaseComObject(excelSheets);
                #endregion
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
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

                string excelName = baseFilePath + "\\XLT\\FabricIronTest.xltx";
                string[] columnNames = new string[]
                {
                "Roll", "Dyelot", "HorizontalOriginal", "VerticalOriginal", "Result", "HorizontalTest1", "HorizontalTest2", "HorizontalTest3", "HorizontalAverage", "HorizontalRate",
                "VerticalTest1", "VerticalTest2", "VerticalTest3", "VerticalAverage", "VerticalRate", "InspDate", "Inspector", "Name", "Remark", "LastUpdate",
                };

                var ret = Array.CreateInstance(typeof(object), dtIronDetail.Rows.Count, columnNames.Length) as object[,];
                for (int i = 0; i < dtIronDetail.Rows.Count; i++)
                {
                    DataRow row = dtIronDetail.Rows[i];
                    for (int j = 0; j < columnNames.Length; j++)
                    {
                        ret[i, j] = row[columnNames[j]];
                    }
                }

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
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

                // 撈seasonID
                // 撈取seasonID
                List<Orders> listOrders = _OrdersProvider.Get(new Orders() { ID = fabricCrkShrkTestIron_Main.POID }).ToList();

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
                excelSheets.Cells[2, 2] = fabricCrkShrkTestIron_Main.ReportNo;
                excelSheets.Cells[3, 2] = fabricCrkShrkTestIron_Main.POID;
                excelSheets.Cells[3, 4] = fabricCrkShrkTestIron_Main.SEQ;
                excelSheets.Cells[3, 6] = fabricCrkShrkTestIron_Main.ColorID;
                excelSheets.Cells[3, 8] = fabricCrkShrkTestIron_Main.StyleID;
                excelSheets.Cells[3, 10] = seasonID;
                excelSheets.Cells[4, 2] = fabricCrkShrkTestIron_Main.SCIRefno;
                excelSheets.Cells[4, 4] = fabricCrkShrkTestIron_Main.ExportID;
                excelSheets.Cells[4, 6] = fabricCrkShrkTestIron_Main.Iron;
                excelSheets.Cells[4, 8] = fabricCrkShrkTestIron_Main.IronDate == null ? string.Empty : ((DateTime)fabricCrkShrkTestIron_Main.IronDate).ToString("yyyy/MM/dd");
                excelSheets.Cells[4, 10] = fabricCrkShrkTestIron_Main.BrandID;
                excelSheets.Cells[5, 2] = fabricCrkShrkTestIron_Main.Refno;
                excelSheets.Cells[5, 4] = fabricCrkShrkTestIron_Main.ArriveQty;
                excelSheets.Cells[5, 6] = fabricCrkShrkTestIron_Main.WhseArrival == null ? string.Empty : ((DateTime)fabricCrkShrkTestIron_Main.WhseArrival).ToString("yyyy/MM/dd");
                excelSheets.Cells[5, 8] = fabricCrkShrkTestIron_Main.Supp;
                excelSheets.Cells[5, 10] = fabricCrkShrkTestIron_Main.NonIron.ToString();

                int defaultRowCount = 10;
                int otherCount = dtIronDetail.Rows.Count - defaultRowCount;
                if (otherCount > 0)
                {
                    // Rate List 複製Row：若有第4筆，則複製一次；有第5筆，則複製2次
                    for (int i = 0; i < otherCount; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range paste = excelSheets.get_Range($"A{defaultRowCount + i}", Type.Missing);
                        Microsoft.Office.Interop.Excel.Range copyRow = excelSheets.get_Range($"A{defaultRowCount}").EntireRow;
                        paste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }
                }

                int RowIdx = 0;
                foreach (DataRow dr in dtIronDetail.Rows)
                {
                    int colIndex = 1;
                    foreach (string col in columnNames)
                    {
                        excel.Cells[RowIdx + 7, colIndex] = dtIronDetail.Rows[RowIdx][col].ToString();
                        colIndex++;
                    }

                    RowIdx++;
                }

                #region 添加圖片
                Excel.Range cellBeforePicture = excelSheets.Cells[22 + otherCount, 1];
                if (fabricCrkShrkTestIron_Main.IronTestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(fabricCrkShrkTestIron_Main.IronTestBeforePicture, fabricCrkShrkTestIron_Main.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    excelSheets.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 2, 323, 255);
                }

                Excel.Range cellAfterPicture = excelSheets.Cells[22 + otherCount, 7];
                if (fabricCrkShrkTestIron_Main.IronTestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(fabricCrkShrkTestIron_Main.IronTestAfterPicture, fabricCrkShrkTestIron_Main.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    excelSheets.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 2, 323, 255);
                }
                #endregion

                excel.Cells.EntireColumn.AutoFit();    // 自動欄寬
                excel.Cells.EntireRow.AutoFit();       ////自動欄高

                #region Save & Show Excel

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }
                char[] invalidChars = Path.GetInvalidFileNameChars();

                foreach (char invalidChar in invalidChars)
                {
                    tmpName = tmpName.Replace(invalidChar.ToString(), "");
                }

                string filexlsx = tmpName + ".xlsx";
                string fileNamePDF = tmpName + ".pdf";

                string filepathpdf = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF);
                string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);


                Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(filepath);

                excelFileName = filexlsx;

                if (IsPDF)
                {
                    if (ConvertToPDF.ExcelToPDF(filepath, filepathpdf))
                    {
                        excelFileName = fileNamePDF;
                        result.Result = true;
                    }
                    else
                    {
                        result.ErrorMessage = "Convert To PDF Fail";
                        result.Result = false;
                    }

                }

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                Marshal.ReleaseComObject(excelSheets);
                #endregion
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

                tmpName = $"Fabric Wash Test_{fabricCrkShrkTestWash_Main.POID}_" +
                        $"{fabricCrkShrkTestWash_Main.StyleID}_" +
                        $"{fabricCrkShrkTestWash_Main.Refno}_" +
                        $"{fabricCrkShrkTestWash_Main.ColorID}_" +
                        $"{fabricCrkShrkTestWash_Main.Wash}_" +
                        $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

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
                if (testCode.Any())
                {
                    excelSheets.Cells[1, 1] = $@"Quality Wash Test Report({testCode.FirstOrDefault().TestCode})";
                }
                excelSheets.Cells[2, 2] = fabricCrkShrkTestWash_Main.ReportNo;
                excelSheets.Cells[3, 2] = fabricCrkShrkTestWash_Main.POID;
                excelSheets.Cells[3, 4] = fabricCrkShrkTestWash_Main.SEQ;
                excelSheets.Cells[3, 6] = fabricCrkShrkTestWash_Main.ColorID;
                excelSheets.Cells[3, 8] = fabricCrkShrkTestWash_Main.StyleID;
                excelSheets.Cells[3, 10] = seasonID;
                excelSheets.Cells[4, 2] = fabricCrkShrkTestWash_Main.SCIRefno;
                excelSheets.Cells[4, 4] = fabricCrkShrkTestWash_Main.ExportID;
                excelSheets.Cells[4, 6] = fabricCrkShrkTestWash_Main.Wash;
                excelSheets.Cells[4, 8] = fabricCrkShrkTestWash_Main.WashDate == null ? string.Empty : ((DateTime)fabricCrkShrkTestWash_Main.WashDate).ToString("yyyy/MM/dd");
                excelSheets.Cells[4, 10] = fabricCrkShrkTestWash_Main.BrandID;
                excelSheets.Cells[5, 2] = fabricCrkShrkTestWash_Main.Refno;
                excelSheets.Cells[5, 4] = fabricCrkShrkTestWash_Main.ArriveQty;
                excelSheets.Cells[5, 6] = fabricCrkShrkTestWash_Main.WhseArrival == null ? string.Empty : ((DateTime)fabricCrkShrkTestWash_Main.WhseArrival).ToString("yyyy/MM/dd");
                excelSheets.Cells[5, 8] = fabricCrkShrkTestWash_Main.Supp;
                excelSheets.Cells[5, 10] = fabricCrkShrkTestWash_Main.NonWash;

                int defaultRowCount = 24;
                int otherCount = dtWashDetail.Rows.Count - defaultRowCount;
                if (otherCount > 0)
                {
                    // Rate List 複製Row：若有第4筆，則複製一次；有第5筆，則複製2次
                    for (int i = 0; i < otherCount; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range paste = excelSheets.get_Range($"A{defaultRowCount + i}", Type.Missing);
                        Microsoft.Office.Interop.Excel.Range copyRow = excelSheets.get_Range($"A{defaultRowCount}").EntireRow;
                        paste.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, copyRow.Copy(Type.Missing));
                    }
                }

                int RowIdx = 0;
                foreach (DataRow dr in dtWashDetail.Rows)
                {
                    int colIndex = 1;
                    foreach (string col in columnNames)
                    {
                        excelSheets.Cells[RowIdx + 7, colIndex] = dtWashDetail.Rows[RowIdx][col].ToString();
                        colIndex++;
                    }

                    RowIdx++;
                }

                // SkewnessOption欄位名稱改變
                switch (skewnessOption)
                {
                    case "1":
                        excelSheets.Cells[6, 16] = "AC";
                        excelSheets.Cells[6, 17] = "BD";
                        break;
                    case "2":
                        excelSheets.Cells[6, 16] = "AA’";
                        excelSheets.Cells[6, 17] = "DD’";
                        excelSheets.Cells[6, 18] = "AB";
                        excelSheets.Cells[6, 19] = "CD";
                        break;
                    case "3":
                        excelSheets.Cells[6, 16] = "AA’";
                        excelSheets.Cells[6, 17] = "AB";
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

                #region 添加圖片
                Excel.Range cellBeforePicture = excelSheets.Cells[34 + otherCount, 1];
                if (fabricCrkShrkTestWash_Main.WashTestBeforePicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(fabricCrkShrkTestWash_Main.WashTestBeforePicture, fabricCrkShrkTestWash_Main.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    excelSheets.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellBeforePicture.Left + 2, cellBeforePicture.Top + 20, 323, 255);
                }

                Excel.Range cellAfterPicture = excelSheets.Cells[34 + otherCount, 7];
                if (fabricCrkShrkTestWash_Main.WashTestAfterPicture != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(fabricCrkShrkTestWash_Main.WashTestAfterPicture, fabricCrkShrkTestWash_Main.ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    excelSheets.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, cellAfterPicture.Left + 2, cellAfterPicture.Top + 20, 323, 255);
                }
                #endregion

                #region Save & Show 

                if (!string.IsNullOrWhiteSpace(AssignedFineName))
                {
                    tmpName = AssignedFineName;
                }

                char[] invalidChars = Path.GetInvalidFileNameChars();

                foreach (char invalidChar in invalidChars)
                {
                    tmpName = tmpName.Replace(invalidChar.ToString(), "");
                }
                string filexlsx = tmpName + ".xlsx";
                string fileNamePDF = tmpName + ".pdf";

                string filepathpdf = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF);
                string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);


                Excel.Workbook workbook = excel.ActiveWorkbook;
                workbook.SaveAs(filepath);

                excelFileName = filexlsx;

                if (IsPDF)
                {
                    if (ConvertToPDF.ExcelToPDF(filepath, filepathpdf))
                    {
                        excelFileName = fileNamePDF;
                        result.Result = true;
                    }
                    else
                    {
                        result.ErrorMessage = "Convert To PDF Fail";
                        result.Result = false;
                    }

                }

                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(excel);
                Marshal.ReleaseComObject(excelSheets);
                #endregion
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.Replace("'", string.Empty);
            }

            return result;
        }
        public BaseResult ToReport_Crocking(long ID, bool IsPDF, out string excelFileName, string AssignedFineName = "")
        {
            BaseResult result = new BaseResult();
            _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
            _QualityBrandTestCodeProvider = new QualityBrandTestCodeProvider(Common.ManufacturingExecutionDataAccessLayer);
            List<Crocking_Excel> dataList_Head = new List<Crocking_Excel>();
            List<Crocking_Excel> dataList_Body = new List<Crocking_Excel>();
            excelFileName = string.Empty;
            string tmpName = string.Empty;

            dataList_Head = _FabricCrkShrkTestProvider.CrockingTest_ToExcel_Head(ID).ToList();
            dataList_Body = _FabricCrkShrkTestProvider.CrockingTest_ToExcel_Body(ID).ToList();

            if (!dataList_Head.Any())
            {
                result.Result = false;
                result.ErrorMessage = "Data not found!";
                return result;
            }

            var testCode = _QualityBrandTestCodeProvider.Get(dataList_Head.FirstOrDefault().BrandID, "Fabric Crocking & Shrinkage Test-Crocking");

            tmpName = $"Fabric Crocking Test_{dataList_Head.FirstOrDefault().POID}_" +
                    $"{dataList_Head.FirstOrDefault().StyleID}_" +
                    $"{dataList_Head.FirstOrDefault().Refno}_" +
                    $"{dataList_Head.FirstOrDefault().Color}_" +
                    $"{dataList_Head.FirstOrDefault().Crocking}_" +
                    $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";

            string basefileName = "FabricCrockingTest";

            switch (dataList_Head.FirstOrDefault().BrandID.ToUpper())
            {
                case "ADIDAS":
                    basefileName = "FabricCrockingTest_ByBrand_A";
                    break;
                case "U.ARMOUR":
                    basefileName = "FabricCrockingTest_ByBrand_A";
                    break;
                case "NIKE":
                    basefileName = "FabricCrockingTest_ByBrand_A";
                    break;
                case "GYMSHARK":
                    basefileName = "FabricCrockingTest_ByBrand_A";
                    break;
                case "LLL":
                    basefileName = "FabricCrockingTest_ByBrand_B";
                    break;
                case "N.FACE":
                    basefileName = "FabricCrockingTest_ByBrand_B";
                    break;
                case "REI":
                    basefileName = "FabricCrockingTest_ByBrand_B";
                    break;
                default:
                    basefileName = "FabricCrockingTest";
                    break;
            }

            string openfilepath = System.Web.HttpContext.Current.Server.MapPath("~/") + $"XLT\\{basefileName}.xltx";

            Microsoft.Office.Interop.Excel.Application excel = MyUtility.Excel.ConnectExcel(openfilepath);
            excel.DisplayAlerts = false; // 設定Excel的警告視窗是否彈出
            //excel.Visible = true;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = excel.ActiveWorkbook.Worksheets[1]; // 取得工作表

            if (testCode.Any())
            {
                worksheet.Cells[1, 1] = $@"Crocking Fastness Test Report({testCode.FirstOrDefault().TestCode})";
            }
            worksheet.Cells[2, 3] = dataList_Head.FirstOrDefault().ReportNo;

            worksheet.Cells[3, 3] = dataList_Head.FirstOrDefault().SubmitDate.HasValue ? dataList_Head.FirstOrDefault().SubmitDate.Value.ToString("yyyy/MM/dd") : string.Empty;
            worksheet.Cells[3, 7] = dataList_Head.FirstOrDefault().SubmitDate;

            worksheet.Cells[4, 3] = dataList_Head.FirstOrDefault().SeasonID;
            worksheet.Cells[4, 7] = dataList_Head.FirstOrDefault().BrandID;

            worksheet.Cells[5, 3] = dataList_Head.FirstOrDefault().StyleID;
            worksheet.Cells[5, 7] = dataList_Head.FirstOrDefault().POID;

            worksheet.Cells[6, 3] = dataList_Head.FirstOrDefault().Article;

            worksheet.Cells[7, 3] = dataList_Head.FirstOrDefault().SCIRefno_Color;
            worksheet.Cells[8, 7] = dataList_Head.FirstOrDefault().Color;

            worksheet.Cells[16, 2] = dataList_Head.FirstOrDefault().Remark;

            worksheet.Cells[72, 3] = dataList_Head.FirstOrDefault().Inspector;

            #region 根據品牌調整圖片、簽名

            if (basefileName == "FabricCrockingTest")
            {
                worksheet.Cells[72, 3] = dataList_Head.FirstOrDefault().Inspector;

                Excel.Range CrockingTestPicture1 = worksheet.Cells[18, 1];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture1 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture1, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture1.Left + 2, CrockingTestPicture1.Top + 2, 323, 255);
                }

                Excel.Range CrockingTestPicture2 = worksheet.Cells[18, 5];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture2 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture2, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture2.Left + 2, CrockingTestPicture2.Top + 2, 323, 255);
                }
                Excel.Range CrockingTestPicture3 = worksheet.Cells[46, 1];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture3 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture3, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture3.Left + 2, CrockingTestPicture3.Top + 2, 323, 255);
                }

                Excel.Range CrockingTestPicture4 = worksheet.Cells[46, 5];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture4 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture4, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture4.Left + 2, CrockingTestPicture4.Top + 2, 323, 255);
                }
            }
            else if (basefileName == "FabricCrockingTest_ByBrand_A")
            {
                worksheet.Cells[72, 3] = dataList_Head.FirstOrDefault().Inspector;

                Excel.Range CrockingTestPicture1 = worksheet.Cells[18, 1];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture1 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture1, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture1.Left + 2, CrockingTestPicture1.Top + 2, 323, 255);
                }

                Excel.Range CrockingTestPicture2 = worksheet.Cells[18, 5];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture2 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture2, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture2.Left + 2, CrockingTestPicture2.Top + 2, 323, 255);
                }
                Excel.Range CrockingTestPicture3 = worksheet.Cells[46, 1];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture3 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture3, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture3.Left + 2, CrockingTestPicture3.Top + 2, 323, 255);
                }

                Excel.Range CrockingTestPicture4 = worksheet.Cells[46, 5];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture4 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture4, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture4.Left + 2, CrockingTestPicture4.Top + 2, 323, 255);
                }
            }
            else if (basefileName == "FabricCrockingTest_ByBrand_B")
            {
                worksheet.Cells[46, 3] = dataList_Head.FirstOrDefault().Inspector;

                Excel.Range CrockingTestPicture1 = worksheet.Cells[18, 1];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture1 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture1, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture1.Left + 2, CrockingTestPicture1.Top + 2, 323, 255);
                }

                Excel.Range CrockingTestPicture2 = worksheet.Cells[18, 5];
                if (dataList_Head.FirstOrDefault().CrockingTestPicture2 != null)
                {
                    string imgPath = ToolKit.PublicClass.AddImageSignWord(dataList_Head.FirstOrDefault().CrockingTestPicture2, dataList_Head.FirstOrDefault().ReportNo, ToolKit.PublicClass.SingLocation.MiddleItalic);
                    worksheet.Shapes.AddPicture(imgPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, CrockingTestPicture2.Left + 2, CrockingTestPicture2.Top + 2, 323, 255);
                }
            }
            #endregion

            // 複製Row：表身幾筆，就幾個Row
            if (dataList_Body.Any() && dataList_Body.Count > 1)
            {
                int copyCtn = dataList_Body.Count - 1;
                Excel.Range paste = worksheet.get_Range($"A13:A13", Type.Missing).EntireRow;
                Excel.Range copyRange = worksheet.get_Range("A13:H14").EntireRow;
                for (int j = 1; j <= copyCtn; j++)
                {
                    paste.Insert(Excel.XlInsertShiftDirection.xlShiftDown, copyRange.Copy(Type.Missing));
                }
            }

            int ctn = 0;
            foreach (var item in dataList_Body)
            {
                int rowIdx = ctn * 2;
                worksheet.Cells[13 + rowIdx, 1] = item.Roll;
                worksheet.Cells[13 + rowIdx, 2] = item.Dyelot;

                worksheet.Cells[13 + rowIdx, 4] = item.DryScale;
                worksheet.Cells[13 + rowIdx, 5] = item.DryScale_Weft;
                worksheet.Cells[13 + rowIdx, 6] = item.WetScale;
                worksheet.Cells[13 + rowIdx, 7] = item.WetScale_Weft;

                worksheet.Cells[14 + rowIdx, 4] = item.ResultDry;
                worksheet.Cells[14 + rowIdx, 5] = item.ResultDry_Weft;
                worksheet.Cells[14 + rowIdx, 6] = item.ResultWet;
                worksheet.Cells[14 + rowIdx, 7] = item.ResultWet_Weft;
                ctn++;
            }

            #region Save & Show Excel

            if (!string.IsNullOrWhiteSpace(AssignedFineName))
            {
                tmpName = AssignedFineName;
            }
            char[] invalidChars = Path.GetInvalidFileNameChars();

            foreach (char invalidChar in invalidChars)
            {
                tmpName = tmpName.Replace(invalidChar.ToString(), "");
            }

            string filexlsx = tmpName + ".xlsx";
            string fileNamePDF = tmpName + ".pdf";

            string filepathpdf = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileNamePDF);
            string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", filexlsx);


            Excel.Workbook workbook = excel.ActiveWorkbook;
            workbook.SaveAs(filepath);

            excelFileName = filexlsx;

            if (IsPDF)
            {
                if (ConvertToPDF.ExcelToPDF(filepath, filepathpdf))
                {
                    excelFileName = fileNamePDF;
                    result.Result = true;
                }
                else
                {
                    result.ErrorMessage = "Convert To PDF Fail";
                    result.Result = false;
                }

            }

            workbook.Close();
            excel.Quit();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excel);


            result.Result = true;
            #endregion

            return result;
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
