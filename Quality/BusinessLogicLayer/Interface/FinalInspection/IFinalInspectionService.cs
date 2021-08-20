using DatabaseObject;
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

        FinalInspection GetFinalInspection(string FinalInspectionID);
    }
}
