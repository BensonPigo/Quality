using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentDefectTypeProvider
    {
        IList<GarmentDefectType> Get();
    }
}
