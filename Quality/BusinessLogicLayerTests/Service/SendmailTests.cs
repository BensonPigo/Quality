using BusinessLogicLayer.Interface;
using DatabaseObject.RequestModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class SendMailTests
    {
        [TestMethod()]
        public void Sendmail()
        {
            try
            {
                ISendMailService _RFTPerLineService = new SendMailService();
                SendMail_Request request = new SendMail_Request()
                {
                    To= "jeff.yeh@sportscity.com.tw",
                    CC="jeff.yeh@sportscity.com.tw",
                    Subject="oo",
                    Body="xx",
                };
               var result = _RFTPerLineService.SendMail(request);
                Assert.IsTrue(result.result);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}