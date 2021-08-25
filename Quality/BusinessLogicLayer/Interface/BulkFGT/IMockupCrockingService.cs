using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IMockupCrockingService
    {
        MockupCrocking_ViewModel GetMockupCrocking(MockupCrocking MockupCrocking);

        MockupUpdatePicture_Result UpdatePicture(MockupCrocking MockupCrocking);
    }
}
