using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface ISewingLineProvider
    {
        IList<SewingLine> GetSewinglineID();
    }
}
