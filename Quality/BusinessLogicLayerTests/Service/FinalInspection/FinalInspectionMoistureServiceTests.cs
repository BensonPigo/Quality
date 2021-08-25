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

        [TestMethod()]
        public void GetMoistureForInspectionTest()
        {
            try
            {
                IFinalInspectionMoistureService finalInspectionMoistureService = new FinalInspectionMoistureService();

                Moisture result = finalInspectionMoistureService.GetMoistureForInspection("ESPCH21080001");

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

        [TestMethod()]
        public void UpdateMoistureTest()
        {
            try
            {
                IFinalInspectionMoistureService finalInspectionMoistureService = new FinalInspectionMoistureService();

                MoistureResult moistureResult = new MoistureResult()
                {
                    FinalInspectionID = "ESPCH21080001",
                    Article = "048675",
                    FinalInspection_OrderCartonUkey = 192,
                    Instrument = "Aqua Boy",
                    Fabrication = "100% Cotton",
                    GarmentTop = 50,
                    GarmentMiddle = 60,
                    GarmentBottom = 15,
                    CTNInside = 17,
                    CTNOutside = 18,
                    Action = "Others",
                    Result = "Y",
                    Remark = "UpdateMoistureTest",
                    AddName = "SCIMIS",
                };

                BaseResult result = finalInspectionMoistureService.UpdateMoistureBySave(moistureResult);

                if (!result && result.ErrorMessage != "Moisture data already exists, please use View to delete first")
                {
                    Assert.Fail(result.ErrorMessage);
                }

                result = finalInspectionMoistureService.UpdateMoistureByNext(moistureResult);

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

        [TestMethod()]
        public void DeleteMoistureTest()
        {
            try
            {
                IFinalInspectionMoistureService finalInspectionMoistureService = new FinalInspectionMoistureService();

                List<ViewMoistureResult> listViewMoistureResult = finalInspectionMoistureService.GetViewMoistureResult("ESPCH21080001");

                BaseResult result = finalInspectionMoistureService.DeleteMoisture(listViewMoistureResult[0].Ukey);

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