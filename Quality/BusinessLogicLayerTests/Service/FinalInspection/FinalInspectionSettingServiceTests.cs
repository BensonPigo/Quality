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
        private List<string> listOrders = new List<string>()
                {
                    "21010007IR",
                    "21010007IR001"
                };

        private string POID = "21010007IR";

        private string factory = "ESP";

        [TestMethod()]
        public void GetSettingForInspectionTest()
        {
            try
            {
                IFinalInspectionSettingService finalInspectionSettingService = new FinalInspectionSettingService();

                

                Setting result = finalInspectionSettingService.GetSettingForInspection(this.POID, this.listOrders, factory);

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

                Setting setting = finalInspectionSettingService.GetSettingForInspection(this.POID, this.listOrders, factory);

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