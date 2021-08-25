using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel.FinalInspection;
using DatabaseObject;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class FinalInspectionMeasurementServiceTests
    {
        [TestMethod()]
        public void GetMeasurementForInspectionTest()
        {
            try
            {
                IFinalInspectionMeasurementService finalInspectionMeasurementService = new FinalInspectionMeasurementService();

                Measurement result = finalInspectionMeasurementService.GetMeasurementForInspection("ESPCH21080001", "SCIMIS");

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetMeasurementViewItemTest()
        {
            try
            {
                IFinalInspectionMeasurementService finalInspectionMeasurementService = new FinalInspectionMeasurementService();

                List<MeasurementViewItem> result = finalInspectionMeasurementService.GetMeasurementViewItem("ESPCH21080001");

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void UpdateMeasurementTest()
        {
            try
            {
                IFinalInspectionMeasurementService finalInspectionMeasurementService = new FinalInspectionMeasurementService();

                Measurement measurement = finalInspectionMeasurementService.GetMeasurementForInspection("ESPCH21080001", "SCIMIS");

                measurement.ListMeasurementItem[10].ResultSizeSpec = "3";
                measurement.SelectedSize = measurement.ListMeasurementItem[10].Size;
                measurement.SelectedArticle = "048675";
                measurement.SelectedProductType = "T";

                BaseResult result = finalInspectionMeasurementService.UpdateMeasurement(measurement, "SCIMIS");

                if (!result)
                {
                    Assert.Fail(result.ErrorMessage);
                }

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}