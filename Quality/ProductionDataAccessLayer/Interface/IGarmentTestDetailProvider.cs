using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestDetailProvider
    {
        IList<GarmentTest_Detail_ViewModel> Get_GarmentTestDetail(GarmentTest filter);

        IList<Order_Qty> GetSizeCode(string OrderID, string Article);
    }
}
