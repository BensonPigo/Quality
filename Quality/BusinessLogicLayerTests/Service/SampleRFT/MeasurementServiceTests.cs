using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service.SampleRFT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface.SampleRFT;

namespace BusinessLogicLayer.Service.SampleRFT.Tests
{
    [TestClass()]
    public class MeasurementServiceTests
    {
        [TestMethod()]
        public void MeasurementGetParaTest()
        {
            try
            {
                IMeasurementService service = new MeasurementService();
                var result = service.MeasurementGetPara("20030480II");
                Assert.IsTrue(result.Articles.Count > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
            
        }
    }
}