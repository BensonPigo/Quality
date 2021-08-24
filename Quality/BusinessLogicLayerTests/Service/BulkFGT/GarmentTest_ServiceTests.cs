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
                var query = _IGarmentTestProvider.Get_GarmentTest(
                    new GarmentTest_ViewModel
                    {
                        StyleID = "NF0A3SR4",
                        BrandID = "N.FACE",
                        Article = "N8E",
                        SeasonID = "19FW",
                        MDivisionid = "Vm2",
                    });

                result.garmentTest = query.FirstOrDefault();

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
    }
}