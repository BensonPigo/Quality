using ADOHelper.Utility;
using BusinessLogicLayer.Interface;
using DatabaseObject.RequestModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
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
                List<string> x = new List<string> { "C:\\Git\\Quality\\Quality\\BusinessLogicLayerTests\\bin\\Debug\\TMP\\MockupOven20210907dae2b1e1-64d6-4478-ba8c-a605f2e73731.pdf" };
                SendMail_Request request = new SendMail_Request()
                {
                    To= "jeff.yeh@sportscity.com.tw",
                    CC="jeff.yeh@sportscity.com.tw",
                    Subject="oo",
                    Body="xx",
                    FileonServer = x,

                };
               var result = MailTools.SendMail(request);
                Assert.IsTrue(result.result);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}