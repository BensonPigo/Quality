using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface
{
    public interface IQueryService
    {
        BaseResult SendMail(string finalInspectionID, string WebHost, bool isTest);

        List<QueryFinalInspection> GetFinalinspectionQueryList_Default(QueryFinalInspection_ViewModel request);
        List<QueryFinalInspection> GetFinalinspectionQueryList(QueryFinalInspection_ViewModel request);

        QueryReport GetFinalInspectionReport(string finalInspectionID);
        List<FinalInspectionSignature> GetSignatureUserList(string UserIDs);
        BaseResult InsertFinalInspectionSignatureUser(string FinalInspectionID, string JobTitle, List<FinalInspectionSignature> allData);
        QueryReport GetFinalInspectionSignature(FinalInspectionSignature Req);
    }
}
