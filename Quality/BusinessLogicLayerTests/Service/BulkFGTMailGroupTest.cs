using BusinessLogicLayer.Interface.BulkFGT;
using BusinessLogicLayer.Service.BulkFGT;
using DatabaseObject.ManufacturingExecutionDB;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace BusinessLogicLayerTests.Service
{
    [TestClass]
    public class BulkFGTMailGroupTest
    {
        [TestMethod]
        public void MailGroupGetTest()
        {
            try
            {
                IBulkFGTMailGroup_Service _BulkFGTMailGroup_Service = new BulkFGTMailGroup_Service();
                Quality_MailGroup quality_Mail = new Quality_MailGroup() { Type = "BulkFGT" };
                var _Para = _BulkFGTMailGroup_Service.MailGroupGet(quality_Mail);
                Assert.IsTrue(_Para[0].ToAddress == "Dyson.chen@sportscity.com.tw");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
        [TestMethod]
        public void MailGroupSave()
        {
            try
            {
                IBulkFGTMailGroup_Service _BulkFGTMailGroup_Service = new BulkFGTMailGroup_Service();
                Quality_MailGroup quality_Mail = new Quality_MailGroup() { FactoryID = "ESP",Type = "BulkFGT", GroupName = "aGroup", ToAddress = "aaa99@aa.aa.aa", CcAddress = "bb99b@bb.bb.bb" };
                var _Para = _BulkFGTMailGroup_Service.MailGroupSave(quality_Mail, 0);
                Assert.IsTrue(_Para.Result && _Para.ErrMsg == null);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}
