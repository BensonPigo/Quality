using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class CFTCommentsTests
    {
        [TestMethod()]
        public void GetRFT_OrderComments()
        {
            try
            {
                ICFTCommentsService _CFTCommentsService = new CFTCommentsService();
                CFTComments_where CFTComments = new CFTComments_where() { OrderID = "20120739GGS01"};
                var _Para = _CFTCommentsService.GetRFT_OrderComments(CFTComments);
                Assert.IsTrue(_Para.Count == 1 );
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}