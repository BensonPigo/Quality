using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class AuthorityServiceTests
    {
        [TestMethod()]
        public void GetAlUserTest()
        {
            try
            {
                IAuthorityService _AuthorityService = new AuthorityService();
                var result = _AuthorityService.GetAlUser("esp");

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}