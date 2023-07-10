using DatabaseObject;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionMeasurementService
    {
        ServiceMeasurement GetMeasurementForInspection(string finalInspectionID, string userID);

        BaseResult UpdateMeasurement(ServiceMeasurement measurement, string userID);

        List<MeasurementViewItem> GetMeasurementViewItem(string finalInspectionID);
    }
}
