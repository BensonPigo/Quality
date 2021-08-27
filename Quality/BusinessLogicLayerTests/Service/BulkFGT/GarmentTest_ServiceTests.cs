using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service.BulkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using DatabaseObject.ViewModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;

namespace BusinessLogicLayer.Service.BulkFGT.Tests
{
    [TestClass()]
    public class GarmentTest_ServiceTests
    {
        [TestMethod()]
        public void GetGarmentTestTest()
        {
            try
            {
                IGarmentTestProvider _IGarmentTestProvider = new GarmentTestProvider(Common.ProductionDataAccessLayer);
                IGarmentTestDetailProvider _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
                GarmentTest_Result result = new GarmentTest_Result();
                //var query = _IGarmentTestProvider.Get_GarmentTest(
                //    new GarmentTest_ViewModel
                //    {
                //        StyleID = "NF0A3SR4",
                //        BrandID = "N.FACE",
                //        Article = "N8E",
                //        SeasonID = "19FW",
                //        MDivisionid = "Vm2",
                //    });

                //result.garmentTest = query.FirstOrDefault();

                // Detail
                result.garmentTest_Details = _IGarmentTestDetailProvider.Get_GarmentTestDetail(
                    new GarmentTest_ViewModel
                    {
                        ID = result.garmentTest.ID
                    }).ToList();


                Assert.IsTrue(result.garmentTest_Details.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_SizeCodeTest()
        {
            try
            {
                List<string> result = new List<string>();
                IGarmentTestDetailProvider _IGarmentTestDetailProvider = new GarmentTestDetailProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailProvider.GetSizeCode("ESPLO14120030", "010-BLAK").Select(x => x.SizeCode).ToList();

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_ShrinkageTest()
        {
            try
            {
                IList<GarmentTest_Detail_Shrinkage> result = new List<GarmentTest_Detail_Shrinkage>();
                IGarmentTestDetailShrinkageProvider _IGarmentTestDetailShrinkageProvider = new GarmentTestDetailShrinkageProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailShrinkageProvider.Get_GarmentTest_Detail_Shrinkage(9244, "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_SpiralityTest()
        {
            try
            {
                IList<Garment_Detail_Spirality> result = new List<Garment_Detail_Spirality>();
                IGarmentDetailSpiralityProvider _IGarmentDetailSpiralityProvider = new GarmentDetailSpiralityProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentDetailSpiralityProvider.Get_Garment_Detail_Spirality(16608, "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_ApperanceTest()
        {
            try
            {
                IList<GarmentTest_Detail_Apperance_ViewModel> result = new List<GarmentTest_Detail_Apperance_ViewModel>();
                IGarmentTestDetailApperanceProvider _IGarmentTestDetailApperanceProvider = new GarmentTestDetailApperanceProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailApperanceProvider.Get_GarmentTest_Detail_Apperance(18578, "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_FGWTTest()
        {
            try
            {
                IList<GarmentTest_Detail_FGWT_ViewModel> result = new List<GarmentTest_Detail_FGWT_ViewModel>();
                IGarmentTestDetailFGWTProvider _IGarmentTestDetailFGWTProvider = new GarmentTestDetailFGWTProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailFGWTProvider.Get_GarmentTest_Detail_FGWT(16615, "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }

        [TestMethod()]
        public void Get_FGPTTest()
        {
            try
            {
                IList<GarmentTest_Detail_FGPT_ViewModel> result = new List<GarmentTest_Detail_FGPT_ViewModel>();
                IGarmentTestDetailFGPTProvider _IGarmentTestDetailFGPTProvider = new GarmentTestDetailFGPTProvider(Common.ProductionDataAccessLayer);
                result = _IGarmentTestDetailFGPTProvider.Get_GarmentTest_Detail_FGPT(16608, "1");

                Assert.IsTrue(result.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }
    }
}