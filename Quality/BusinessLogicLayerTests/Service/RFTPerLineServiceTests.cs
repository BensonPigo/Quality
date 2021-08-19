using BusinessLogicLayer.Interface;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class RFTPerLineServiceTests
    {
        [TestMethod()]
        public void GetQueryParaTest()
        {
            try
            {
                IRFTPerLineService _RFTPerLineService = new RFTPerLineService();
                var _Para = _RFTPerLineService.GetQueryPara();
                Assert.IsTrue(_Para.Months.Count == 12 && _Para.Years.Min() == "2021" && _Para.Years.Max() == "2021");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }


        [TestMethod()]
        public void RFTPerLineQueryTest()
        {
            try
            {
                IRFTPerLineService _RFTPerLineService = new RFTPerLineService();
                var data = _RFTPerLineService.RFTPerLineQuery("es2", "2021", "8");
                Assert.IsTrue(data.dailyRFTs[0].Month == "August" && data.monthlyRFTs[0].Month == "August");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}