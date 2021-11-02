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
using System.Drawing;

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
                return _FabricCrkShrkTestProvider.GetTestPOID(string.Empty);
            }
        }

        private long ID
        {
            get
            {
                IFabricCrkShrkTestProvider _FabricCrkShrkTestProvider = new FabricCrkShrkTestProvider(Common.ProductionDataAccessLayer);
                return _FabricCrkShrkTestProvider.GetTestFIRID(string.Empty);
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
                fabricCrkShrkTest_Result = fabricCrkShrkTest_Service.GetFabricCrkShrkTest_Result("21041716WW");

                fabricCrkShrkTest_Result.Main.FirLaboratoryRemark = "995";

                fabricCrkShrkTest_Result.Details[0].ReceiveSampleDate = DateTime.Now;

                fabricCrkShrkTest_Result.Details[0].ReceiveSampleDate = DateTime.Now;
                fabricCrkShrkTest_Result.Details[0].NonCrocking = false;
                fabricCrkShrkTest_Result.Details[0].NonHeat = false;
                fabricCrkShrkTest_Result.Details[0].NonWash = false;
                fabricCrkShrkTest_Result.Details[0].Wash = "";
                fabricCrkShrkTest_Result.Details[0].Heat = "";
                fabricCrkShrkTest_Result.Details[0].Crocking = "";

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
                Bitmap bitmap = new Bitmap(Image.FromFile(@"TestResource\001.jpg"));
                ImageConverter converter = new ImageConverter();
                byte[] testImgByte = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));

                bitmap = new Bitmap(Image.FromFile(@"TestResource\Koala.jpg"));
                byte[] testImgByte2 = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));
                fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestBeforePicture = testImgByte;
                fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestAfterPicture = testImgByte2;

                fabricCrkShrkTestCrocking_Result.Crocking_Detail[0].DryScale = "995";
                fabricCrkShrkTestCrocking_Result.Crocking_Detail[0].ResultDry = "Pass";

                BaseResult baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestCrockingDetail(fabricCrkShrkTestCrocking_Result, "SCIMIS");

                FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result2 = fabricCrkShrkTest_Service.GetFabricCrkShrkTestCrocking_Result(testID);
                fabricCrkShrkTestCrocking_Result.Crocking_Detail.Add(fabricCrkShrkTestCrocking_Result2.Crocking_Detail[0]);
                fabricCrkShrkTestCrocking_Result.Crocking_Detail[fabricCrkShrkTestCrocking_Result.Crocking_Detail.Count - 1].Roll = "995";

                baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestCrockingDetail(fabricCrkShrkTestCrocking_Result, "SCIMIS");
                if (!baseResult)
                {
                    Assert.Fail(baseResult.ErrorMessage);
                }

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

        [TestMethod()]
        public void ToPdfFabricCrkShrkTestCrockingDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                string excelName;
                BaseResult baseResult = fabricCrkShrkTest_Service.ToPdfFabricCrkShrkTestCrockingDetail(360043, out excelName, true);

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
        public void GetFabricCrkShrkTestHeat_ResultTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                FabricCrkShrkTestHeat_Result result;
                result = fabricCrkShrkTest_Service.GetFabricCrkShrkTestHeat_Result(this.ID);
                fabricCrkShrkTest_Service.GetFabricCrkShrkTestHeat_Result(-1);

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
        public void SaveFabricCrkShrkTestHeatDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                long testID = this.ID;

                FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result;
                fabricCrkShrkTestHeat_Result = fabricCrkShrkTest_Service.GetFabricCrkShrkTestHeat_Result(testID);

                fabricCrkShrkTestHeat_Result.Heat_Detail[0].HorizontalTest1 = 0.02M;
                fabricCrkShrkTestHeat_Result.Heat_Detail[0].HorizontalTest2 = 0.01M;
                fabricCrkShrkTestHeat_Result.Heat_Detail[0].HorizontalTest3 = 0.02M;
                fabricCrkShrkTestHeat_Result.Heat_Detail[0].VerticalTest1 = 0.02M;
                fabricCrkShrkTestHeat_Result.Heat_Detail[0].VerticalTest2 = 0.02M;
                fabricCrkShrkTestHeat_Result.Heat_Detail[0].VerticalTest3 = -0.01M;
                fabricCrkShrkTestHeat_Result.Heat_Detail[0].Remark = "9527";
                Bitmap bitmap = new Bitmap(Image.FromFile(@"TestResource\001.jpg"));
                ImageConverter converter = new ImageConverter();
                byte[] testImgByte = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));

                bitmap = new Bitmap(Image.FromFile(@"TestResource\Koala.jpg"));
                byte[] testImgByte2 = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));
                fabricCrkShrkTestHeat_Result.Heat_Main.HeatTestBeforePicture = testImgByte;
                fabricCrkShrkTestHeat_Result.Heat_Main.HeatTestAfterPicture = testImgByte2;

                BaseResult baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestHeatDetail(fabricCrkShrkTestHeat_Result, "SCIMIS");

                FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result2 = fabricCrkShrkTest_Service.GetFabricCrkShrkTestHeat_Result(testID);
                fabricCrkShrkTestHeat_Result.Heat_Detail.Add(fabricCrkShrkTestHeat_Result2.Heat_Detail[0]);
                fabricCrkShrkTestHeat_Result.Heat_Detail[fabricCrkShrkTestHeat_Result.Heat_Detail.Count - 1].Roll = "995";

                baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestHeatDetail(fabricCrkShrkTestHeat_Result, "SCIMIS");

                fabricCrkShrkTestHeat_Result.Heat_Detail.RemoveAt(0);

                baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestHeatDetail(fabricCrkShrkTestHeat_Result, "SCIMIS");

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
        public void EncodeFabricCrkShrkTestHeatDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                string testResult;
                BaseResult baseResult = fabricCrkShrkTest_Service.EncodeFabricCrkShrkTestHeatDetail(this.ID, "SCIMIS", out testResult);

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
        public void SendHeatFailResultMailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                SendMail_Result result = fabricCrkShrkTest_Service.SendHeatFailResultMail("aaron.shie@sportscity.com.tw", "aaron.shie@sportscity.com.tw", this.ID, true);

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
        public void AmendFabricCrkShrkTestHeatDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                BaseResult baseResult = fabricCrkShrkTest_Service.AmendFabricCrkShrkTestHeatDetail(this.ID);

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
        public void ToExcelFabricCrkShrkTestHeatDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                string excelName;
                BaseResult baseResult = fabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestHeatDetail(this.ID, out excelName, true);

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
        public void GetScaleIDsTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                List<string> result = fabricCrkShrkTest_Service.GetScaleIDs();

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetFabricCrkShrkTestWash_ResultTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                FabricCrkShrkTestWash_Result result;
                result = fabricCrkShrkTest_Service.GetFabricCrkShrkTestWash_Result(this.ID);
                fabricCrkShrkTest_Service.GetFabricCrkShrkTestWash_Result(-1);

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
        public void SaveFabricCrkShrkTestWashDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                long testID = this.ID;

                FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result;
                fabricCrkShrkTestWash_Result = fabricCrkShrkTest_Service.GetFabricCrkShrkTestWash_Result(testID);

                fabricCrkShrkTestWash_Result.Wash_Detail[0].HorizontalTest1 = 0.02M;
                fabricCrkShrkTestWash_Result.Wash_Detail[0].HorizontalTest2 = 0.01M;
                fabricCrkShrkTestWash_Result.Wash_Detail[0].HorizontalTest3 = 0.02M;
                fabricCrkShrkTestWash_Result.Wash_Detail[0].VerticalTest1 = 0.02M;
                fabricCrkShrkTestWash_Result.Wash_Detail[0].VerticalTest2 = 0.02M;
                fabricCrkShrkTestWash_Result.Wash_Detail[0].VerticalTest3 = -0.01M;
                fabricCrkShrkTestWash_Result.Wash_Detail[0].Remark = "9527";
                fabricCrkShrkTestWash_Result.Wash_Main.SkewnessOptionID = "3";
                fabricCrkShrkTestWash_Result.Wash_Detail[0].SkewnessTest1 = 5;
                fabricCrkShrkTestWash_Result.Wash_Detail[0].SkewnessTest2 = 3;
                fabricCrkShrkTestWash_Result.Wash_Detail[0].SkewnessTest3 = 6;
                fabricCrkShrkTestWash_Result.Wash_Detail[0].SkewnessTest4 = 2;
                Bitmap bitmap = new Bitmap(Image.FromFile(@"TestResource\001.jpg"));
                ImageConverter converter = new ImageConverter();
                byte[] testImgByte = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));

                bitmap = new Bitmap(Image.FromFile(@"TestResource\Koala.jpg"));
                byte[] testImgByte2 = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));
                fabricCrkShrkTestWash_Result.Wash_Main.WashTestBeforePicture = testImgByte;
                fabricCrkShrkTestWash_Result.Wash_Main.WashTestAfterPicture = testImgByte2;

                BaseResult baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestWashDetail(fabricCrkShrkTestWash_Result, "SCIMIS");

                FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result2 = fabricCrkShrkTest_Service.GetFabricCrkShrkTestWash_Result(testID);
                fabricCrkShrkTestWash_Result.Wash_Detail.Add(fabricCrkShrkTestWash_Result2.Wash_Detail[0]);
                fabricCrkShrkTestWash_Result.Wash_Detail[fabricCrkShrkTestWash_Result.Wash_Detail.Count - 1].Roll = "995";

                baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestWashDetail(fabricCrkShrkTestWash_Result, "SCIMIS");

                fabricCrkShrkTestWash_Result.Wash_Detail.RemoveAt(0);

                baseResult = fabricCrkShrkTest_Service.SaveFabricCrkShrkTestWashDetail(fabricCrkShrkTestWash_Result, "SCIMIS");

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
        public void EncodeFabricCrkShrkTestWashDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                string testResult;
                BaseResult baseResult = fabricCrkShrkTest_Service.EncodeFabricCrkShrkTestWashDetail(this.ID, "SCIMIS", out testResult);

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
        public void SendWashFailResultMailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                SendMail_Result result = fabricCrkShrkTest_Service.SendWashFailResultMail("aaron.shie@sportscity.com.tw", "aaron.shie@sportscity.com.tw", this.ID, true);

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
        public void AmendFabricCrkShrkTestWashDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();

                BaseResult baseResult = fabricCrkShrkTest_Service.AmendFabricCrkShrkTestWashDetail(this.ID);

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
        public void ToExcelFabricCrkShrkTestWashDetailTest()
        {
            try
            {
                IFabricCrkShrkTest_Service fabricCrkShrkTest_Service = new FabricCrkShrkTest_Service();
                string excelName;
                BaseResult baseResult = fabricCrkShrkTest_Service.ToExcelFabricCrkShrkTestWashDetail(364797, out excelName, true);

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