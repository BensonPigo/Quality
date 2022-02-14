using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Interface.SampleRFT
{
    public interface IMeasurementService
    {
        Measurement_Request MeasurementGetPara(string OrderID, string FactoryID);

        Measurement_ResultModel MeasurementGet(Measurement_Request measurement);

        Measurement_Request DeleteMeasurementImage(long ID);
        Measurement_Request MeasurementToExcel(string OrderID, string FactoryID, bool test = false);
    }
}
