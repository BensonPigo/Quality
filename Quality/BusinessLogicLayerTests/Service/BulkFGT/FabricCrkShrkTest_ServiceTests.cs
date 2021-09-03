using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface;
using DatabaseObject.ResultModel;
using DatabaseObject;
using ADOHelper.Template.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class FabricCrkShrkTest_ServiceTests
    {
        private string POID
        {
            get
            {
                IFabricCrkShrkTestProvider _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                return _FabricCrkShrkTestProvider.GetTestPOID();
            }
        }

        private long ID
        {
            get
            {
                IFabricCrkShrkTestProvider _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                return _FabricCrkShrkTestProvider.GetTestFIRID();
            }
        }

        [TestMethod()]
        public void GetFabricCrkShrkTest_ResultTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                FabricCrkShrkTest_Result fabricCrkShrkTest_Result;
                fabricCrkShrkTest_Result = fabricCrkShrkTest_Service.GetFabricCrkShrkTest_Result(this.POID);
                fabricCrkShrkTest_Service.GetFabricCrkShrkTest_Result("9527");

            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("data not found"))
                {
                    Assert.IsTrue(true);
                    return;
                }

                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void SaveFabricCrkShrkTestMainTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                FabricCrkShrkTest_Result fabricCrkShrkTest_Result;
                fabricCrkShrkTest_Result = fabricCrkShrkTest_Service.GetFabricCrkShrkTest_Result(this.POID);

                fabricCrkShrkTest_Result.Main.FirLaboratoryRemark = "995";

                fabricCrkShrkTest_Result.Details[0].ReceiveSampleDate = DateTime.Now;

                fabricCrkShrkTest_Result.Details[1].ReceiveSampleDate = DateTime.Now;
                fabricCrkShrkTest_Result.Details[1].NonCrocking = true;
                fabricCrkShrkTest_Result.Details[1].NonHeat = true;
                fabricCrkShrkTest_Result.Details[1].NonWash = false;

                BaseResult baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestMain(fabricCrkShrkTest_Result);

                if (!baseResult)
                {
                    Assert.Fail(baseResult.ErrorMessage);
                }

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetFabricCrkShrkTestCrocking_ResultTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                FabricCrkShrkTestCrocking_Result result;
                result = fabricCrkShrkTest_Service.GetFabricCrkShrkTestCrocking_Result(this.ID);
                fabricCrkShrkTest_Service.GetFabricCrkShrkTestCrocking_Result(-1);

            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("No data found"))
                {
                    Assert.IsTrue(true);
                    return;
                }

                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void SaveFabricCrkShrkTestCrockingDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                long testID = this.ID;

                FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result;
                fabricCrkShrkTestCrocking_Result = fabricCrkShrkTest_Service.GetFabricCrkShrkTestCrocking_Result(testID);

                fabricCrkShrkTestCrocking_Result.Crocking_Detail[0].DryScale = "995";
                fabricCrkShrkTestCrocking_Result.Crocking_Detail[0].ResultDry = "Pass";

                BaseResult baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestCrockingDetail(fabricCrkShrkTestCrocking_Result, "SCIMIS");

                FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result2 = fabricCrkShrkTest_Service.GetFabricCrkShrkTestCrocking_Result(testID);
                fabricCrkShrkTestCrocking_Result.Crocking_Detail.Add(fabricCrkShrkTestCrocking_Result2.Crocking_Detail[0]);
                fabricCrkShrkTestCrocking_Result.Crocking_Detail[fabricCrkShrkTestCrocking_Result.Crocking_Detail.Count - 1].Roll = "995";

                baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestCrockingDetail(fabricCrkShrkTestCrocking_Result, "SCIMIS");

                fabricCrkShrkTestCrocking_Result.Crocking_Detail.RemoveAt(0);

                baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestCrockingDetail(fabricCrkShrkTestCrocking_Result, "SCIMIS");

                if (!baseResult)
                {
                    Assert.Fail(baseResult.ErrorMessage);
                }

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void EncodeFabricCrkShrkTestCrockingDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                string testResult;
                BaseResult baseResult = fabricCrkShrkTest_Service.EncodeFabricCrkShrkTestCrockingDetail(this.ID, "SCIMIS", out testResult);

                if (!baseResult)
                {
                    Assert.Fail(baseResult.ErrorMessage);
                }

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void SendCrockingFailResultMailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                SendMail_Result result = fabricCrkShrkTest_Service.SendCrockingFailResultMail("aaron.shie@sportscity.com.tw", "aaron.shie@sportscity.com.tw", this.ID, true);

                if (!result.result)
                {
                    Assert.Fail(result.resultMsg);
                }

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void AmendFabricCrkShrkTestCrockingDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                BaseResult baseResult = fabricCrkShrkTest_Service.AmendFabricCrkShrkTestCrockingDetail(this.ID);

                if (!baseResult)
                {
                    Assert.Fail(baseResult.ErrorMessage);
                }

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void ToExcelFabricCrkShrkTestCrockingDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                string excelName;
                BaseResult baseResult = fabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestCrockingDetail(this.ID, out excelName, true);

                if (!baseResult)
                {
                    Assert.Fail(baseResult.ErrorMessage);
                }

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}