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
                    FileList = new System.Collections.Generic.List<string>() { "C:\\Git\\Quality\\Quality\\BusinessLogicLayerTests\\bin\\Debug\\TMP\\MockupCrocking20210830a3c71190-59d7-41c6-93e6-d1d5f7651789.pdf" },
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