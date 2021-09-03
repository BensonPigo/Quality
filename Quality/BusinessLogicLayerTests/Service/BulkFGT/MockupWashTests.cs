using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using DatabaseObject.ProductionDB;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject.RequestModel;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class MockupWashTests
    {
        [TestMethod()]
        public void GetMockupWash()
        {
            try
            {
                IMockupWashService _MockupWashService = new MockupWashService();
                MockupWash_Request MockupWash = new MockupWash_Request () 
                { 
                    ReportNo = "PHWA201000482",
                    //Type = "B",
                    //BrandID = "ADIDAS",
                    //SeasonID= "20SS", 
                    //StyleID= "S201CSPM108",
                    //Article= "FL0237"
                };
                var mockupWash = _MockupWashService.GetMockupWash(MockupWash);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetAccessoryRefNo()
        {
            try
            {
                IMockupWashService _MockupWashService = new MockupWashService();
                AccessoryRefNo_Request accessoryRefNo_Request = new AccessoryRefNo_Request()
                {
                    BrandID = "U.ARMOUR",
                    SeasonID = "20FW",
                    StyleID = "1342962",
                    //StyleUkey =75468,
                };
                var selectListItems = _MockupWashService.GetAccessoryRefNo(accessoryRefNo_Request);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetPDF()
        {
            try
            {
                IMockupWashService _MockupWashService = new MockupWashService();
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}