using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.ResultModel;
using DatabaseObject;
using System.Drawing;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class FabricOvenTestServiceTests
    {
        [TestMethod()]
        public void GetFabricOvenTest_ResultTest()
        {
            try
            {
                IFabricOvenTestService fabricOvenTestService = new FabricOvenTestService();
                FabricOvenTest_Result fabricOvenTest_Result;
                fabricOvenTest_Result = fabricOvenTestService.GetFabricOvenTest_Result("21051739BB");
                fabricOvenTestService.GetFabricOvenTest_Result("9527");

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
        public void SaveFabricOvenTestMainTest()
        {
            try
            {
                IFabricOvenTestService fabricOvenTestService = new FabricOvenTestService();

                FabricOvenTest_Main fabricOvenTest_Main = new FabricOvenTest_Main()
                {
                    POID = "21051739BB",
                    Remark = "9527"
                };

                BaseResult baseResult = fabricOvenTestService.SaveFabricOvenTestMain(fabricOvenTest_Main);

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
        public void GetFabricOvenTest_Detail_ResultTest()
        {
            try
            {
                IFabricOvenTestService fabricOvenTestService = new FabricOvenTestService();

                FabricOvenTest_Result fabricOvenTest_Result;
                fabricOvenTest_Result = fabricOvenTestService.GetFabricOvenTest_Result("21051739BB");

                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result;
                fabricOvenTest_Detail_Result = fabricOvenTestService.GetFabricOvenTest_Detail_Result(fabricOvenTest_Result.Main.POID, fabricOvenTest_Result.Details[0].TestNo);

                fabricOvenTest_Detail_Result = fabricOvenTestService.GetFabricOvenTest_Detail_Result(fabricOvenTest_Result.Main.POID, "");

                fabricOvenTest_Detail_Result = fabricOvenTestService.GetFabricOvenTest_Detail_Result(fabricOvenTest_Result.Main.POID, "55");

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
        public void SaveFabricOvenTestDetailTest()
        {
            try
            {
                IFabricOvenTestService fabricOvenTestService = new FabricOvenTestService();

                FabricOvenTest_Result fabricOvenTest_Result;
                fabricOvenTest_Result = fabricOvenTestService.GetFabricOvenTest_Result("21051739BB");

                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result = fabricOvenTestService.GetFabricOvenTest_Detail_Result(fabricOvenTest_Result.Main.POID, fabricOvenTest_Result.Details[0].TestNo);

                fabricOvenTest_Detail_Result.Main.TestNo = string.Empty;

                BaseResult baseResult = fabricOvenTestService.SaveFabricOvenTestDetail(fabricOvenTest_Detail_Result, "SCIMIS");

                fabricOvenTest_Detail_Result.Main.TestNo = "1";
                Bitmap bitmap = new Bitmap(Image.FromFile(@"TestResource\001.jpg"));
                ImageConverter converter = new ImageConverter();
                byte[] testImgByte = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));
                fabricOvenTest_Detail_Result.Main.TestBeforePicture = testImgByte;

                testImgByte = (byte[])converter.ConvertTo(new Bitmap(Image.FromFile(@"TestResource\Koala.jpg")), typeof(byte[]));
                fabricOvenTest_Detail_Result.Main.TestAfterPicture = null;

                fabricOvenTest_Detail_Result.Details[0].Remark = "9527";
                fabricOvenTest_Detail_Result.Details[0].SubmitDate = DateTime.Now.Date;
                fabricOvenTest_Detail_Result.Details[0].Roll = "3";
                fabricOvenTest_Detail_Result.Details[0].ChangeScale = "5";
                fabricOvenTest_Detail_Result.Details[0].StainingScale = "6";
                fabricOvenTest_Detail_Result.Details[0].ResultChange = "test";
                fabricOvenTest_Detail_Result.Details[0].ResultStain = "test";
                fabricOvenTest_Detail_Result.Details[0].Temperature = "87";
                fabricOvenTest_Detail_Result.Details[0].Time = "7";

                baseResult = fabricOvenTestService.SaveFabricOvenTestDetail(fabricOvenTest_Detail_Result, "SCIMIS");

                FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result2 = fabricOvenTestService.GetFabricOvenTest_Detail_Result(fabricOvenTest_Result.Main.POID, fabricOvenTest_Result.Details[0].TestNo);

                fabricOvenTest_Detail_Result.Details.Add(fabricOvenTest_Detail_Result2.Details[0]);
                fabricOvenTest_Detail_Result.Details[fabricOvenTest_Detail_Result.Details.Count - 1].SEQ = "99-95";
                baseResult = fabricOvenTestService.SaveFabricOvenTestDetail(fabricOvenTest_Detail_Result, "SCIMIS");

                fabricOvenTest_Detail_Result.Details.RemoveAt(fabricOvenTest_Detail_Result.Details.Count - 1);
                baseResult = fabricOvenTestService.SaveFabricOvenTestDetail(fabricOvenTest_Detail_Result, "SCIMIS");

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
        public void EncodeFabricOvenTestDetailTest()
        {
            try
            {
                IFabricOvenTestService fabricOvenTestService = new FabricOvenTestService();
                string ovenTestResult;
                BaseResult baseResult = fabricOvenTestService.EncodeFabricOvenTestDetail("21051739BB", "3", out ovenTestResult);

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
        public void AmendFabricOvenTestDetailTest()
        {
            try
            {
                IFabricOvenTestService fabricOvenTestService = new FabricOvenTestService();

                BaseResult baseResult = fabricOvenTestService.AmendFabricOvenTestDetail("21051739BB", "3");

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
        public void SendFailResultMailTest()
        {
            try
            {
                IFabricOvenTestService fabricOvenTestService = new FabricOvenTestService();

                SendMail_Result result = fabricOvenTestService.SendFailResultMail("aaron.shie@sportscity.com.tw", "aaron.shie@sportscity.com.tw", "21051739BB", "3", true);

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
        public void ToExcelFabricOvenTestDetailTest()
        {
            try
            {
                IFabricOvenTestService fabricOvenTestService = new FabricOvenTestService();
                string excelName;
                BaseResult baseResult = fabricOvenTestService.ToExcelFabricOvenTestDetail("21051739BB", "1", out excelName, true);

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
        public void ToPdfFabricOvenTestDetailTest()
        {
            try
            {
                IFabricOvenTestService fabricOvenTestService = new FabricOvenTestService();
                string excelName;
                BaseResult baseResult = fabricOvenTestService.ToPdfFabricOvenTestDetail("21051739BB", "1", out excelName, true);

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