using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System.Collections.Generic;
using System.Data;
using System.Web.Mvc;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IMeasurementProvider
    {
        IList<Order_Qty> GetAtricle(string OrderID);

        IList<Measurement_Request> Get_OrdersPara(string OrderID, string FactoryID);

        int Get_Total_Measured_Qty();

        int Get_Measured_Qty(Measurement_Request measurement);

        DataTable Get_Measured_Detail(Measurement_Request measurement);

        DataTable Get_CalculateSizeSpec(string diffValue, string Tol);

        IList<Measurement> GetMeasurementsByPOID(string POID, string OrderID, string userID);
        IList<SelectListItem> Get_ImageSource(string OrderID);
        IList<RFT_Inspection_Measurement_Image> Get_ImageList(string OrderID);
        int DeleteMeasurementImgae(long ID);
    }
}
