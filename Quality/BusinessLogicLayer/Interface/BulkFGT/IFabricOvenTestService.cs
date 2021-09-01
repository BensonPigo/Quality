using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IFabricOvenTestService
    {
        FabricOvenTest_Result GetFabricOvenTest_Result(string POID);

        BaseResult SaveFabricOvenTestMain(FabricOvenTest_Main fabricOvenTest_Main);

        FabricOvenTest_Detail_Result GetFabricOvenTest_Detail_Result(string poID, string TestNo);

        BaseResult SaveFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID);

        BaseResult EncodeFabricOvenTestDetail(string poID, string TestNo, out string ovenTestResult);
        BaseResult SendFailResultMail(string toAddress, string ccAddress, string poID, string TestNo);

        BaseResult AmendFabricOvenTestDetail(string poID, string TestNo);

        string ToExcelFabricOvenTestDetail(string poID, string TestNo);

        string ToPdfFabricOvenTestDetail(string poID, string TestNo);
    }
}
