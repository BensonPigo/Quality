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
using System.Drawing;

namespace BusinessLogicLayer.Service.FinalInspectionTests
{
    [TestClass()]
    public class FinalInspectionAddDefectServiceTests
    {
        [TestMethod()]
        public void GetDefectForInspectionTest()
        {
            try
            {
                IFinalInspectionAddDefectService finalInspectionAddDefectService = new FinalInspectionAddDefectService();

                AddDefect result = finalInspectionAddDefectService.GetDefectForInspection("ESPCH21080001");

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
        public void GetDefectImageTest()
        {
            try
            {
                IFinalInspectionAddDefectService finalInspectionAddDefectService = new FinalInspectionAddDefectService();

                List<byte[]> result = finalInspectionAddDefectService.GetDefectImage(289);

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void UpdateFinalInspectionDetailTest()
        {
            try
            {
                IFinalInspectionAddDefectService finalInspectionAddDefectService = new FinalInspectionAddDefectService();

                AddDefect addDefect = finalInspectionAddDefectService.GetDefectForInspection("ESPCH21080001");

                addDefect.RejectQty = 9;
                addDefect.ListFinalInspectionDefectItem[0].Qty = 5;
                addDefect.ListFinalInspectionDefectItem[3].Qty = 6;
                addDefect.InspectionStep = "Insp-CheckList";
                Bitmap bitmap = new Bitmap(Image.FromFile(@"TestResource\001.jpg"));
                ImageConverter converter = new ImageConverter();
                byte[] testImgByte = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));

                addDefect.ListFinalInspectionDefectItem[0].ListFinalInspectionDefectImage.Add(testImgByte);
                BaseResult result = finalInspectionAddDefectService.UpdateFinalInspectionDetail(addDefect, "SCIMIS");

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