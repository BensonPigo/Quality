using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Collections.Generic;
using System.Data;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IMeasurementProvider
    {
        IList<Order_Qty> GetAtricle(string OrderID);

        Measurement_Request Get_OrdersPara(string OrderID);

        int Get_Total_Measured_Qty();

        int Get_Measured_Qty(Measurement_Request measurement);

        DataTable Get_Measured_Detail(Measurement_Request measurement);

        DataTable Get_CalculateSizeSpec(string diffValue, string Tol);
    }
}
