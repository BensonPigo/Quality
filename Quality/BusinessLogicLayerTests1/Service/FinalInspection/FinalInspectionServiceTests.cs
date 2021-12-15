using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class FinalInspectionServiceTests
    {
        [TestMethod()]
        public void GetPivot88JsonTest()
        {
            try
            {
                FinalInspectionService finalInspectionService = new FinalInspectionService();
                string result = JsonConvert.SerializeObject(finalInspectionService.GetPivot88Json("ES3CH21110011"));
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}