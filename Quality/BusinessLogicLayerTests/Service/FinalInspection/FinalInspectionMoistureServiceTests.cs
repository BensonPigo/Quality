using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel.FinalInspection;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class FinalInspectionMoistureServiceTests
    {
        [TestMethod()]
        public void GetViewMoistureResultTest()
        {
            try
            {
                IFinalInspectionMoistureService finalInspectionMoistureService = new FinalInspectionMoistureService();

                List<ViewMoistureResult> result = finalInspectionMoistureService.GetViewMoistureResult("ESPCH21080001");

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}