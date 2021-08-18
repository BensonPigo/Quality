using Microsoft.VisualStudio.TestTools.UnitTesting;
using BusinessLogicLayer.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatabaseObject.ManufacturingExecutionDB;
using BusinessLogicLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using DatabaseObject.ViewModel;
using DatabaseObject.ProductionDB;

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

        [TestMethod()]
        public void RFT_OrderCommentsGetTest()
        {
            try
            {
                RFT_OrderComments rework = new RFT_OrderComments()
                {
                    OrderID = "21061052UC",
                    PMS_RFTCommentsID = "1",
                    Comnments = "test 2021",
                };

                IInspectionService _InspectionService = new InspectionService();
                var _rework = _InspectionService.GetRFT_OrderComments(rework).ToList();
                Assert.IsTrue(_rework.Count > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [TestMethod()]
        public void SaveMeasurementTest()
        {
            try
            {
                IRFTInspectionMeasurementProvider _IRFTInspectionMeasurementProvider = new RFTInspectionMeasurementProvider(Common.ManufacturingExecutionDataAccessLayer);
                IOrdersProvider _IOrdersProvider = new OrdersProvider(Common.ProductionDataAccessLayer);
                IStyleProvider _IStyleProvider = new StyleProvider(Common.ProductionDataAccessLayer);
                List<RFT_Inspection_Measurement_ViewModel> _Inspection_Measurement_ViewModels = new List<RFT_Inspection_Measurement_ViewModel>();

                IList<Orders> ordersList = _IOrdersProvider.Get(new Orders() { ID = "21070015IR003" });
                string strSizeUnit = string.Empty;
                string longStyleUkey = string.Empty;
                if (ordersList.Count > 0)
                {
                    longStyleUkey = ordersList[0].StyleUkey.ToString();
                    IList<Style> StyleList = _IStyleProvider.GetSizeUnit(Convert.ToInt64(longStyleUkey));

                    strSizeUnit = StyleList[0].SizeUnit;
                }

                _Inspection_Measurement_ViewModels = _IRFTInspectionMeasurementProvider.Get(Convert.ToInt64(longStyleUkey), "6", "SCIMIS").ToList();

                int updateCnt = _IRFTInspectionMeasurementProvider.Save(_Inspection_Measurement_ViewModels);
                Assert.IsTrue(updateCnt > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
                throw;
            }

        }

        [TestMethod()]
        public void GetDQSReasonTest()
        {
            try
            {
                IDQSReasonProvider _IDQSReasonProvider = new DQSReasonProvider(Common.ManufacturingExecutionDataAccessLayer);
                List<DQSReason> reasons = new List<DQSReason>();
                reasons = _IDQSReasonProvider.Get(
                 new DQSReason()
                 {
                     Type = "DP",
                 }).ToList();

                Assert.IsTrue(reasons.Count > 0);
            }
            catch (Exception)
            {
                Assert.Fail();
                throw;
            }
        }
    }
}