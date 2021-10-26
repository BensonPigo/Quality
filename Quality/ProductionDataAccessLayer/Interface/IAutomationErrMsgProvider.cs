using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IAutomationErrMsgProvider
    {
        void Insert(AutomationErrMsg automationErrMsg);
    }
}
