using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentDefectCodeProvider
    {
        IList<GarmentDefectCode> Get(GarmentDefectCode Item);
    }
}
