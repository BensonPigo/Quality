using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using System.Collections.Generic;
using System.Web;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IFabricOvenTestService
    {
        FabricOvenTest_Result GetFabricOvenTest_Result(string POID);

        BaseResult SaveFabricOvenTestMain(FabricOvenTest_Main fabricOvenTest_Main);

        BaseResult DeleteOven(string poID, string TestNo);

        FabricOvenTest_Detail_Result GetFabricOvenTest_Detail_Result(string poID, string TestNo ,string BrandID);

        BaseResult SaveFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID);

        BaseResult EncodeFabricOvenTestDetail(string poID, string TestNo, out string ovenTestResult);
        SendMail_Result SendMail(string toAddress, string ccAddress, string poID, string TestNo, bool isTest, string Subject, string Body, List<HttpPostedFileBase> Files);

        BaseResult AmendFabricOvenTestDetail(string poID, string TestNo);

        BaseResult ToExcelFabricOvenTestDetail(string poID, string TestNo, out string excelFileName, bool isTest);

        BaseResult ToPdfFabricOvenTestDetail(string poID, string TestNo, out string pdfFileName, bool isTest, string AssignedFineName = "");
        BaseResult ToPdfFabricOvenTestDetail_Ori(string poID, string TestNo, out string pdfFileName, bool isTest);
    }
}
