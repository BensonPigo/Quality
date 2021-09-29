using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service.FinalInspection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessLogicLayer.Interface;
using DatabaseObject;
using DatabaseObject.ViewModel.FinalInspection;
using DatabaseObject.RequestModel;

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

                BaseResult result = _QueryService.SendMail("ESPCH21080001","XX", true);

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

        [TestMethod()]
        public void GetFinalinspectionQueryListTest()
        {
            try
            {
                IQueryService _QueryService = new QueryService();

                QueryFinalInspection_ViewModel queryFinalInspection = new QueryFinalInspection_ViewModel();

                queryFinalInspection.CustPONO = "21010007IR";

                List<QueryFinalInspection> result = _QueryService.GetFinalinspectionQueryList(queryFinalInspection);

                queryFinalInspection.SP = "21010007IR";
                result = _QueryService.GetFinalinspectionQueryList(queryFinalInspection);

                queryFinalInspection.SciDeliveryStart = DateTime.Parse("2021-01-20");
                result = _QueryService.GetFinalinspectionQueryList(queryFinalInspection);

                queryFinalInspection.InspectionResult = "Pass";
                result = _QueryService.GetFinalinspectionQueryList(queryFinalInspection);

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void GetFinalInspectionReportTest()
        {
            try
            {
                IQueryService _QueryService = new QueryService();

                QueryReport result = _QueryService.GetFinalInspectionReport("ESPCH21080001");

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