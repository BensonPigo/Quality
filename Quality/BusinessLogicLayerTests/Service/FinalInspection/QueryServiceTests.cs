using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service.FinalInspection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface;
using DatabaseObject;

namespace BusinessLogicLayer.Service.FinalInspection.Tests
{
    [TestClass()]
    public class QueryServiceTests
    {
        [TestMethod()]
        public void SendMailTest()
        {
            try
            {
                IQueryService _QueryService = new QueryService();

                BaseResult result = _QueryService.SendMail("ESPCH21080001", true);

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