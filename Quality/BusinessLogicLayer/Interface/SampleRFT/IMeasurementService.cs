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
        Measurement_Request MeasurementGetPara(string OrderID);

        Measurement_ResultModel MeasurementGet(Measurement_Request measurement);

        void MeasurementToExcel(Measurement_Request measurement);
    }
}
