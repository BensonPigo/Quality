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
    public class FinalInspectionSettingServiceTests
    {
        [TestMethod()]
        public void GetSettingForInspectionTest()
        {
            try
            {
                IFinalInspectionSettingService finalInspectionSettingService = new FinalInspectionSettingService();

                List<string> listOrders = new List<string>()
                {
                    "21010007IR",
                    "21010007IR001"
                };

                Setting result = finalInspectionSettingService.GetSettingForInspection("21010007IR", listOrders, "ESP");

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
        public void GetSettingForInspectionTest1()
        {
            try
            {
                IFinalInspectionSettingService finalInspectionSettingService = new FinalInspectionSettingService();

                Setting result = finalInspectionSettingService.GetSettingForInspection("21010007IR");

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
        public void UpdateFinalInspectionTest()
        {
            try
            {
                IFinalInspectionSettingService finalInspectionSettingService = new FinalInspectionSettingService();

                Setting setting = new Setting()
                { 
                    FinalInspectionID = string.Empty,
                    InspectionStage = "",
                    AuditDate = DateTime.Now,
                    SewingLineID = "05",
                    
                };

                //finalInspectionSettingService.UpdateFinalInspection("21010007IR");

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}