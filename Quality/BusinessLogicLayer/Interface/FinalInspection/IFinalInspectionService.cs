using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionService
    {
        IList<Orders> GetOrderForInspection(FinalInspection_Request request);

        DatabaseObject.ManufacturingExecutionDB.FinalInspection GetFinalInspection(string FinalInspectionID);
    }
}
