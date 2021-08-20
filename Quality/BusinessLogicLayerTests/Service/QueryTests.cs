using BusinessLogicLayer.Interface;
using DatabaseObject.ViewModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using DatabaseObject.ProductionDB;
using BusinessLogicLayer.Service.FinalInspection;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class QueryTests
    {
        [TestMethod()]
        public void Get()
        {
            try
            {
                IQueryService _QueryService = new QueryService();
                DatabaseObject.ManufacturingExecutionDB.FinalInspection Req = new DatabaseObject.ManufacturingExecutionDB.FinalInspection();
                var _Para = _QueryService.SendMail(Req);
                Assert.IsTrue(_Para);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}