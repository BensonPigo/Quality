using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Interface.SampleRFT;
using BusinessLogicLayer.Service.SampleRFT;
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
                CFTComments_ViewModel CFTComments = new CFTComments_ViewModel() { OrderID = "20120739GGS01"};
                var _Para = _CFTCommentsService.Get_CFT_OrderComments(CFTComments).DataList;
                Assert.IsTrue(_Para.Count == 1 );
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}