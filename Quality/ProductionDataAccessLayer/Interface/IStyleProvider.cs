using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IStyleProvider
    {
        IList<Style> Get(Style Item);

        IList<Style> GetSizeUnit(Int64 ukey);

        string GetSizeUnitByCustPONO(string CustPONO, string OrderID);

        string GetStyleName(string StyleID, string Season, string Brand);
        string GetStyleCritical(string StyleID, string Season, string Brand);
    }
}
