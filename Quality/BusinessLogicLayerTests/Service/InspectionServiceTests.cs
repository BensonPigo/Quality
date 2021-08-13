using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatabaseObject.ManufacturingExecutionDB;
using BusinessLogicLayer.Interface;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class InspectionServiceTests
    {
        [TestMethod()]
        public void GetReworkCardsTest()
        {
            try
            {
                ReworkCard rework = new ReworkCard()
                {
                    FactoryID = "ES2",
                    Line = "01",
                    Type = "Hard",
                };

                IInspectionService _InspectionService = new InspectionService();
                var _rework = _InspectionService.GetReworkCards(rework).ToList();
                Assert.IsTrue(_rework.Count > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}