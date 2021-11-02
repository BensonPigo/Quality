using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IFIRLaboratoryProvider
    {
        IList<FIR_Laboratory> Get(FIR_Laboratory Item);
    }
}
