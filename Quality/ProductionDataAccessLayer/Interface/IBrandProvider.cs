using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IBrandProvider
    {
        IList<Brand> Get();
    }
}
