using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupCrockingDetailProvider
    {
        IList<MockupCrocking_Detail> Get(MockupCrocking_Detail Item);
    }
}
