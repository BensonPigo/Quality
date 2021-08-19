using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface ISeasonProvider
    {
        IList<Season> Get(Season Item);
    }
}
