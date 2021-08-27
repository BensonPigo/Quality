using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupCrockingService
    {
        MockupCrocking_ViewModel GetMockupCrocking(MockupCrocking MockupCrocking);

        MockupCrocking_ViewModel UpdatePicture(MockupCrocking MockupCrocking);

        MockupCrocking_ViewModel GetExcel(MockupCrocking_ViewModel Req);
    }
}
