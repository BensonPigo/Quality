using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.FinalInspection;
using System.Collections.Generic;
using System.Data;
using System.Web.Mvc;

namespace ProductionDataAccessLayer.Interface
{
    public interface IFabricOvenTestProvider
    {
        FabricOvenTest_Result GetFabricOvenTest_Main(string POID);

        FabricOvenTest_Detail_Result GetFabricOvenTest_Detail(string poID, string TestNo);

        void SaveFabricOvenTestMain(FabricOvenTest_Main fabricOvenTest_Main);

        void AddFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID);

        void EditFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID);

        void EncodeFabricOven(string poID, string TestNo, string result);

        void AmendFabricOven(string poID, string TestNo);

        DataTable GetFailMailContentData(string poID, string TestNo);
    }
}
