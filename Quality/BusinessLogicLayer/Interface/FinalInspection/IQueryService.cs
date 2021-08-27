using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IQueryService
    {
        BaseResult SendMail(string finalInspectionID, bool isTest);

        List<QueryFinalInspection> GetFinalinspectionQueryList(FinalInspection_Query request);

        QueryReport GetFinalInspectionReport(string finalInspectionID);
    }
}
