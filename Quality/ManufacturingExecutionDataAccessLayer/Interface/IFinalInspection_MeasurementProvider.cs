using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IFinalInspection_MeasurementProvider
    {
        IList<QueryReport_Measurement> GetQuery_FinalInspection_Measurement(string ID, string SP);
    }
}
