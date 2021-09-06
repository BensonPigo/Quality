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

namespace BusinessLogicLayer.Service.FinalInspectionTests
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

        private string Mdivision = "VM2";

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

                Setting result = finalInspectionSettingService.GetSettingForInspection("ESPCH21080001");

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
                setting.AcceptQty = 0;
                setting.AcceptableQualityLevelsUkey = "50";
                setting.AuditDate = DateTime.Now;
                setting.InspectionStage = "Final";
                setting.SampleSize = 6;
                setting.SewingLineID = "05";
                setting.SelectCarton[0].Selected = true;
                setting.SelectCarton[1].Selected = true;
                setting.SelectOrderShipSeq[0].Selected = true;
                string result = string.Empty;
                BaseResult baseResult = finalInspectionSettingService.UpdateFinalInspection(setting, "SCIMIS", this.factory, this.Mdivision, out result);
                if (!baseResult)
                {
                    Assert.Fail(baseResult.ErrorMessage);
                }

                setting = finalInspectionSettingService.GetSettingForInspection("ESPCH21080001");
                setting.SampleSize = 0;
                setting.SelectedPO[0].AvailableQty = 20;
                setting.SelectCarton[0].Selected = false;
                setting.SelectCarton[2].Selected = true;
                baseResult = finalInspectionSettingService.UpdateFinalInspection(setting, "SCIMIS", this.factory, this.Mdivision, out result);
                if (!baseResult)
                {
                    Assert.Fail(baseResult.ErrorMessage);
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