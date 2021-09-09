using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service.StyleManagement;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using Sci;

namespace BusinessLogicLayer.Service.StyleManagement.Tests
{
    [TestClass()]
    public class StyleManagementServiceTests
    {
        [TestMethod()]
        public void Get_StyleResult_BrowseTest()
        {
            try
            {
                IStyleManagementService styleManagementService = new StyleManagementService();

                StyleManagement_Request styleManagement_Request = new StyleManagement_Request();

                styleManagement_Request.StyleID = "F2002M203";
                styleManagement_Request.BrandID = "ADIDAS";
                styleManagement_Request.SeasonID = "22SS";

                StyleResult_ViewModel result = styleManagementService.Get_StyleResult_Browse(styleManagement_Request);

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}