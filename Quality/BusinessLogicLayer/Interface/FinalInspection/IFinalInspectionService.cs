using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;

namespace BusinessLogicLayer.Interface
{
    public interface IFinalInspectionService
    {
        Orders Get(FinalInspection_Request request);
    }
}
