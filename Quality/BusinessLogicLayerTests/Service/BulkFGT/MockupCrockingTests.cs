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
    public class MockupCrockingTests
    {
        [TestMethod()]
        public void Get()
        {
            try
            {
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
                MockupCrocking_Request MockupCrocking = new MockupCrocking_Request () { BrandID = "ADIDAS", SeasonID= "19SS", StyleID= "S1953WTR302", Article= "DU1325", Type="S"};
                var mockupCrocking = _MockupCrockingService.GetMockupCrocking(MockupCrocking);
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
                IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}