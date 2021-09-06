using DatabaseObject;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.FinalInspection;
using System;
using System.Collections.Generic;
using System.Data;
using System.Web.Mvc;

namespace ProductionDataAccessLayer.Interface
{
    public interface IFabricCrkShrkTestProvider
    {
        FabricCrkShrkTest_Result GetFabricCrkShrkTest_Main(string POID);
        void SaveFabricCrkShrkTest_Main(FabricCrkShrkTest_Result fabricCrkShrkTest_Result);

        string GetTestPOID();
        long GetTestFIRID();

        // Crocking
        FabricCrkShrkTestCrocking_Main GetFabricCrockingTest_Main(long ID);

        List<FabricCrkShrkTestCrocking_Detail> GetFabricCrockingTest_Detail(long ID);

        int GetCrockingTestOption(long ID);

        void UpdateFabricCrockingTestDetail(FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result, string userID);

        void EncodeFabricCrocking(long ID, string testResult, DateTime? crockingDate, string userID);

        DataTable GetCrockingFailMailContentData(long ID);

        void AmendFabricCrocking(long ID);

        DataTable GetCrockingDetailForReport(long ID);

        DataTable GetCrockingArticleForPdfReport(long ID);
    }
}
