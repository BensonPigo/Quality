using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IMockupCrockingProvider
    {
        IList<MockupCrocking> Get(MockupCrocking Item);

        int Create(MockupCrocking Item);

        int Update(MockupCrocking Item);

        int UpdatePicture(MockupCrocking Item);

        int Delete(MockupCrocking Item);
    }
}
