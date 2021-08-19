using BusinessLogicLayer.Interface.SampleRFT;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.SampleRFT
{
    public class MeasurementService : IMeasurementService
    {
        public Measurement_ResultModel MeasurementGet(Measurement_Request measurement)
        {
            throw new NotImplementedException();
        }

        public Measurement_Request MeasurementGetPara(string OrderID)
        {
            throw new NotImplementedException();
        }

        public void MeasurementToExcel(Measurement_Request measurement)
        {
            throw new NotImplementedException();
        }
    }
}
