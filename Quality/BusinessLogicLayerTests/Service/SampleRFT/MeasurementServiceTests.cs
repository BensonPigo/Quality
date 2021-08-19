using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service.SampleRFT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface.SampleRFT;
using DatabaseObject.RequestModel;

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
                var result = service.MeasurementGetPara("20091188GGS");
                Assert.IsTrue(result.Articles.Count > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }

        }

        [TestMethod()]
        public void MeasurementGetTest()
        {
            try
            {
                IMeasurementService service = new MeasurementService();
                Measurement_Request _Measurement_Request = service.MeasurementGetPara("20091188GGS");
                _Measurement_Request.Factory = "ESP";

                var result = service.MeasurementGet(_Measurement_Request);
                Assert.IsTrue(result.Articles.Count > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void MeasurementGetTest1()
        {
            Assert.Fail();
        }
    }
}