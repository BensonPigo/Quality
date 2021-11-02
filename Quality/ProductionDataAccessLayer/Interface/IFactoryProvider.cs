using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IFactoryProvider
    {
        IList<Factory> GetFtyGroup();

        IList<Factory> GetMDivisionID(string fty);
    }
}
