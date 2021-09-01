using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ResultModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service
{
    public class FabricOvenTestService : IFabricOvenTestService
    {
        IFabricOvenTestProvider _FabricOvenTestProvider;
        IScaleProvider _ScaleProvider;

        public BaseResult AmendFabricOvenTestDetail(string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
            try
            {
                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result = _FabricOvenTestProvider.GetFabricOvenTest_Detail(poID, TestNo);

                if (fabricOvenTest_Detail_Result.Main.Status != "Confirmed")
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Status is {fabricOvenTest_Detail_Result.Main.Status}, can not Amend";
                    return baseResult;
                }

                _FabricOvenTestProvider.AmendFabricOven(poID, TestNo);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public BaseResult EncodeFabricOvenTestDetail(string poID, string TestNo, out string ovenTestResult)
        {
            BaseResult baseResult = new BaseResult();
            _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
            ovenTestResult = string.Empty;
            try
            {
                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result = _FabricOvenTestProvider.GetFabricOvenTest_Detail(poID, TestNo);

                if (fabricOvenTest_Detail_Result.Main.Status != "New")
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Status is {fabricOvenTest_Detail_Result.Main.Status}, can not Encode";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Details.Count == 0)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Data is empty please fill-in data first.";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Details.Any(s => string.IsNullOrEmpty(s.OvenGroup)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Group cannot be empty.";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Details.Any(s => string.IsNullOrEmpty(s.SEQ)))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Seq cannot be empty.";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Details.Any(s => 
                    string.IsNullOrEmpty(s.ChangeScale) ||
                    string.IsNullOrEmpty(s.StainingScale) ||
                    string.IsNullOrEmpty(s.Result) 
                ))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = $"Color Change Scale, Color Staining Scale and Result cannot be empty.";
                    return baseResult;
                }

                string result = fabricOvenTest_Detail_Result.Details.Any(s => s.Result == "Fail") ? "Fail" : "Pass";
                ovenTestResult = result;
                _FabricOvenTestProvider.EncodeFabricOven(poID, TestNo, result);

                return baseResult;
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
                return baseResult;
            }
        }

        public FabricOvenTest_Detail_Result GetFabricOvenTest_Detail_Result(string poID, string TestNo)
        {
            try
            {
                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result = new FabricOvenTest_Detail_Result();
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                _ScaleProvider = new ScaleProvider(Common.ProductionDataAccessLayer);

                fabricOvenTest_Detail_Result = _FabricOvenTestProvider.GetFabricOvenTest_Detail(poID, TestNo);

                fabricOvenTest_Detail_Result.ScaleIDs = _ScaleProvider.Get().Select(s => s.ID).ToList();
                
                return fabricOvenTest_Detail_Result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public FabricOvenTest_Result GetFabricOvenTest_Result(string POID)
        {
            try
            {
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);

                return _FabricOvenTestProvider.GetFabricOvenTest_Main(POID);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public BaseResult SaveFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                if (string.IsNullOrEmpty(fabricOvenTest_Detail_Result.Main.Article))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Article cannot be empty.";
                    return baseResult;
                }

                if (string.IsNullOrEmpty(fabricOvenTest_Detail_Result.Main.Inspector))
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Inspector cannot be empty.";
                    return baseResult;
                }

                if (fabricOvenTest_Detail_Result.Main.InspDate == null)
                {
                    baseResult.Result = false;
                    baseResult.ErrorMessage = "Test Date cannot be empty.";
                    return baseResult;
                }

                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);

                if (string.IsNullOrEmpty(fabricOvenTest_Detail_Result.Main.TestNo))
                {
                    _FabricOvenTestProvider.AddFabricOvenTestDetail(fabricOvenTest_Detail_Result, userID);
                }
                else
                {
                    _FabricOvenTestProvider.EditFabricOvenTestDetail(fabricOvenTest_Detail_Result, userID);
                }
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public BaseResult SaveFabricOvenTestMain(FabricOvenTest_Main fabricOvenTest_Main)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                _FabricOvenTestProvider.SaveFabricOvenTestMain(fabricOvenTest_Main);
            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public BaseResult SendFailResultMail(string toAddress, string ccAddress, string poID, string TestNo)
        {
            BaseResult baseResult = new BaseResult();
            try
            {
                _FabricOvenTestProvider = new FabricOvenTestProvider(Common.ProductionDataAccessLayer);
                DataTable dtResult = _FabricOvenTestProvider.GetFailMailContentData(poID, TestNo);
                string mailBody = MailTools.DataTableChangeHtml(dtResult);
                //string resultMsg = MailTools.MailToHtml(toAddress, );

            }
            catch (Exception ex)
            {
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.ToString();
            }

            return baseResult;
        }

        public string ToExcelFabricOvenTestDetail(string poID, string TestNo)
        {
            throw new NotImplementedException();
        }

        public string ToPdfFabricOvenTestDetail(string poID, string TestNo)
        {
            throw new NotImplementedException();
        }
    }
}
