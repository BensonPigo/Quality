using BusinessLogicLayer.Interface;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using System.Collections.Generic;
using System.Linq;

namespace BusinessLogicLayer.Service
{
    public class MockupCrockingService : IMockupCrockingService
    {
        private IMockupCrockingProvider _MockupCrockingProvider;
        private IMockupCrockingDetailProvider _MockupCrockingDetailProvider;

        public MockupCrocking_ViewModel GetMockupCrocking(MockupCrocking MockupCrocking)
        {
            MockupCrocking_ViewModel MockupCrocking_ViewModel = new MockupCrocking_ViewModel();
            _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
            _MockupCrockingDetailProvider = new MockupCrockingDetailProvider(Common.ProductionDataAccessLayer);
            MockupCrocking_ViewModel.MockupCrocking = _MockupCrockingProvider.Get(MockupCrocking).ToList();
            MockupCrocking_ViewModel.MockupCrocking_Detail = new List<MockupCrocking_Detail>();
            foreach (var item in MockupCrocking_ViewModel.MockupCrocking)
            {
                MockupCrocking_Detail mockupCrocking_Detail = new MockupCrocking_Detail() { ReportNo = item.ReportNo };
                foreach (var MockupCrockingDetail in _MockupCrockingDetailProvider.Get(mockupCrocking_Detail))
                {
                    MockupCrocking_ViewModel.MockupCrocking_Detail.Add(MockupCrockingDetail);
                }
            }

            return MockupCrocking_ViewModel;
        }

        // 單獨更新圖檔欄位
        public MockupUpdatePicture_Result UpdatePicture(MockupCrocking MockupCrocking)
        {
            MockupUpdatePicture_Result updatePicture_Result = new MockupUpdatePicture_Result();
            _MockupCrockingProvider = new MockupCrockingProvider(Common.ProductionDataAccessLayer);
            try
            {
                int updateCT = _MockupCrockingProvider.UpdatePicture(MockupCrocking);
                if (updateCT == 0)
                {
                    updatePicture_Result.result = false;
                    updatePicture_Result.resultMsg = "Not found Crocking Data!";
                }
                else
                {
                    updatePicture_Result.result = true;
                }
            }
            catch (System.Exception ex)
            {
                updatePicture_Result.result = false;
                updatePicture_Result.resultMsg = ex.ToString();
            }

            return updatePicture_Result;
        }
    }
}
