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
                var _Para = _MockupCrockingService.GetMockupCrocking(MockupCrocking);
                Assert.IsTrue(_Para.MockupCrocking.Count > 0 && _Para.ReportNos.Count > 0);
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
                var MockupCrocking_ViewModel = _MockupCrockingService.GetExcel("PHCK180800007");
                string file = MockupCrocking_ViewModel.TempFileName;
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetScale()
        {
            IMockupCrockingService _MockupCrockingService = new MockupCrockingService();
            var x = _MockupCrockingService.GetScale();
            Assert.IsTrue(x.WetScale.Count > 0);
        }
    }
}