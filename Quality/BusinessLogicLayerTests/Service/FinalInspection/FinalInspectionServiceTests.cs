using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using DatabaseObject.ManufacturingExecutionDB;
using ProductionDataAccessLayer.Provider.MSSQL;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;

namespace BusinessLogicLayer.Service.Tests
{
    [TestClass()]
    public class FinalInspectionServiceTests
    {
        [TestMethod()]
        public void GetFinalInspectionTest()
        {
            try
            {
                FinalInspectionService finalInspectionService = new FinalInspectionService();

                DatabaseObject.ManufacturingExecutionDB.FinalInspection result = finalInspectionService.GetFinalInspection("123");

                if (!result.Result)
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

        [TestMethod()]
        public void GetOrderForInspectionTest()
        {
            try
            {
                FinalInspectionService finalInspectionService = new FinalInspectionService();

                FinalInspection_Request inspection_Request = new FinalInspection_Request()
                { 
                    FactoryID = "ESP",
                    StyleID = "LW6BM7S",
                    SP = "21040034IC002",
                    POID = "21040034IC"
                };

                IList<Orders> result = finalInspectionService.GetOrderForInspection(inspection_Request);

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }
}