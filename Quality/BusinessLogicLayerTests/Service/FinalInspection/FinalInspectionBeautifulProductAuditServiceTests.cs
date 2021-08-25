using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel.FinalInspection;
using System.Drawing;
using DatabaseObject;

namespace BusinessLogicLayer.Service.FinalInspectionTests
{
    [TestClass()]
    public class FinalInspectionBeautifulProductAuditServiceTests
    {
        [TestMethod()]
        public void GetBeautifulProductAuditForInspectionTest()
        {
            try
            {
                IFinalInspectionBeautifulProductAuditService finalInspectionBeautifulProductAuditService = new FinalInspectionBeautifulProductAuditService();

                BeautifulProductAudit result = finalInspectionBeautifulProductAuditService.GetBeautifulProductAuditForInspection("ESPCH21080001");

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
        public void GetBACriteriaImageTest()
        {
            try
            {
                IFinalInspectionBeautifulProductAuditService finalInspectionBeautifulProductAuditService = new FinalInspectionBeautifulProductAuditService();

                List<byte[]> result = finalInspectionBeautifulProductAuditService.GetBACriteriaImage(289);

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void UpdateBeautifulProductAuditTest()
        {
            try
            {
                IFinalInspectionBeautifulProductAuditService finalInspectionBeautifulProductAuditService = new FinalInspectionBeautifulProductAuditService();

                BeautifulProductAudit beautifulProductAudit = finalInspectionBeautifulProductAuditService.GetBeautifulProductAuditForInspection("ESPCH21080001");

                beautifulProductAudit.BAQty = 99;
                beautifulProductAudit.ListBACriteria[0].Qty = 5;
                beautifulProductAudit.ListBACriteria[3].Qty = 6;
                beautifulProductAudit.InspectionStep = "Insp-BA";
                Bitmap bitmap = new Bitmap(Image.FromFile(@"TestResource\001.jpg"));
                ImageConverter converter = new ImageConverter();
                byte[] testImgByte = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));

                beautifulProductAudit.ListBACriteria[0].ListBACriteriaImage.Add(testImgByte);
                BaseResult result = finalInspectionBeautifulProductAuditService.UpdateBeautifulProductAudit(beautifulProductAudit, "SCIMIS");

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