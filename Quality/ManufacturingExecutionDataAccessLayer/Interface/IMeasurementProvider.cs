using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IMeasurementProvider
    {
        IList<Order_Qty> GetAtricle(string OrderID);

        Measurement_Request Get_OrdersPara(string OrderID);
    }
}
