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
                Assert.IsTrue(data.monthlyRFTs[0].RFT == (decimal)0.07 &&
                    data.dailyRFTs.Where(w=>w.Line == "01"&& w.Date ==16).Select(s=>s.RFT).FirstOrDefault() == (decimal)0.08);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}