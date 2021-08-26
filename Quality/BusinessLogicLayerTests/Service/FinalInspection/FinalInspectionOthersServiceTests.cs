using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatabaseObject.ViewModel.FinalInspection;
using BusinessLogicLayer.Interface;
using System.Drawing;
using DatabaseObject;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class FinalInspectionOthersServiceTests
    {
        [TestMethod()]
        public void GetOthersForInspectionTest()
        {
            try
            {
                IFinalInspectionOthersService finalInspectionOthersService = new FinalInspectionOthersService();

                Others result = finalInspectionOthersService.GetOthersForInspection("ESPCH21080001");

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
        public void GetOthersImageTest()
        {
            try
            {
                IFinalInspectionOthersService finalInspectionOthersService = new FinalInspectionOthersService();

                List<byte[]> result = finalInspectionOthersService.GetOthersImage("ESPCH21080001");

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void UpdateOthersBackTest()
        {
            try
            {
                IFinalInspectionOthersService finalInspectionOthersService = new FinalInspectionOthersService();

                Others others = finalInspectionOthersService.GetOthersForInspection("ESPCH21080001");

                others.OthersRemark = "995995";

                Bitmap bitmap = new Bitmap(Image.FromFile(@"TestResource\001.jpg"));
                ImageConverter converter = new ImageConverter();
                byte[] testImgByte = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));

                others.ListOthersImageItem.Add(testImgByte);
                BaseResult result = finalInspectionOthersService.UpdateOthersBack(others, "SCIMIS");

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
        public void UpdateOthersSubmitTest()
        {
            try
            {
                IFinalInspectionOthersService finalInspectionOthersService = new FinalInspectionOthersService();

                Others others = finalInspectionOthersService.GetOthersForInspection("ESPCH21080001");

                others.OthersRemark = "995995";

                Bitmap bitmap = new Bitmap(Image.FromFile(@"TestResource\001.jpg"));
                ImageConverter converter = new ImageConverter();
                byte[] testImgByte = (byte[])converter.ConvertTo(bitmap, typeof(byte[]));

                others.ListOthersImageItem.Add(testImgByte);
                BaseResult result = finalInspectionOthersService.UpdateOthersSubmit(others, "SCIMIS");

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