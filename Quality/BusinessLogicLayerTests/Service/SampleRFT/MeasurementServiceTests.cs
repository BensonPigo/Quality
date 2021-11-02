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
                var result = service.MeasurementGetPara("20091188GGS", "ESP");
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
                Measurement_Request _Measurement_Request = service.MeasurementGetPara("20091188GGS", "ESP");
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

        [TestMethod()]
        public void MeasurementToExcelTest()
        {
            try
            {
                Measurement_Request all_Data = new Measurement_Request();
                IMeasurementService _Service = new MeasurementService();
                all_Data = _Service.MeasurementToExcel("21031564GGS02", "ES2", true);
                Assert.IsTrue(all_Data.Result);
            }
            catch (Exception ex)
            {
                Assert.Fail();
                throw ex;
            }

        }
    }
}