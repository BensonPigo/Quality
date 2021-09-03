using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IStyleBOAProvider
    {
        IList<Style_BOA> GetAccessoryRefNo(AccessoryRefNo_Request Item);
    }
}
