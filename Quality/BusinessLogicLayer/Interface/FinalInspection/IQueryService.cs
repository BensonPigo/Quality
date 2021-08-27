using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;

namespace BusinessLogicLayer.Interface
{
    public interface IQueryService
    {
        BaseResult SendMail(string finalInspectionID, bool isTest);
    }
}
