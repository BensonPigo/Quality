using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IRFTInspectionMeasurementProvider
    {
        IList<RFT_Inspection_Measurement_ViewModel> Get(Int64 StyleUkey, string SizeCode, string UserID);
    }
}
