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
        string GetFactoryNameEN(string POID, string FactoryID);
        FabricOvenTest_Result GetFabricOvenTest_Main(string POID);

        FabricOvenTest_Detail_Result GetFabricOvenTest_Detail(string poID, string TestNo, string BrandID ="");

        void SaveFabricOvenTestMain(FabricOvenTest_Main fabricOvenTest_Main);

        void AddFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID, out string TestNo);

        void EditFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID);

        void EncodeFabricOven(string poID, string TestNo, string result);

        void AmendFabricOven(string poID, string TestNo);

        DataTable GetFailMailContentData(string poID, string TestNo);

        DataTable GetOvenDetailForExcel(string poID, string TestNo);

        DataTable GetOven(string poID, string TestNo);

        void DeleteOven(string poID, string TestNo);

    }
}
