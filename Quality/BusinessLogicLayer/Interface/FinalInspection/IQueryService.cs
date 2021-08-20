using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;

namespace BusinessLogicLayer.Interface
{
    public interface IQueryService
    {
        bool SendMail(DatabaseObject.ManufacturingExecutionDB.FinalInspection Req);
    }
}
