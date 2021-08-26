using DatabaseObject;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionMeasurementService
    {
        Measurement GetMeasurementForInspection(string finalInspectionID, string userID);

        BaseResult UpdateMeasurement(Measurement measurement, string userID);

        List<MeasurementViewItem> GetMeasurementViewItem(string finalInspectionID);
    }
}
