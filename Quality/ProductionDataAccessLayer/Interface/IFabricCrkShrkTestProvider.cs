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
        string GetFactoryNameEN(string factory);
        FabricCrkShrkTest_Result GetFabricCrkShrkTest_Main(string POID);
        void SaveFabricCrkShrkTest_Main(FabricCrkShrkTest_Result fabricCrkShrkTest_Result);

        string GetTestPOID(string where);
        long GetTestFIRID(string where);

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
        
        List<Crocking_Excel> CrockingTest_ToExcel(long ID);
        List<Crocking_Excel> CrockingTest_ToExcel_Head(long ID);
        List<Crocking_Excel> CrockingTest_ToExcel_Body(long ID);

        // Heat
        FabricCrkShrkTestHeat_Main GetFabricHeatTest_Main(long ID);

        List<FabricCrkShrkTestHeat_Detail> GetFabricHeatTest_Detail(long ID);

        void UpdateFabricHeatTestDetail(FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result, string userID);

        void EncodeFabricHeat(long ID, string testResult, DateTime? heatDate, string userID);

        DataTable GetHeatFailMailContentData(long ID);

        void AmendFabricHeat(long ID);

        DataTable GetHeatDetailForReport(long ID);


        // Iron
        FabricCrkShrkTestIron_Main GetFabricIronTest_Main(long ID);

        List<FabricCrkShrkTestIron_Detail> GetFabricIronTest_Detail(long ID);

        void UpdateFabricIronTestDetail(FabricCrkShrkTestIron_Result fabricCrkShrkTestIron_Result, string userID);

        void EncodeFabricIron(long ID, string testResult, DateTime? IronDate, string userID);

        DataTable GetIronFailMailContentData(long ID);

        void AmendFabricIron(long ID);

        DataTable GetIronDetailForReport(long ID);

        // Wash
        FabricCrkShrkTestWash_Main GetFabricWashTest_Main(long ID);

        List<FabricCrkShrkTestWash_Detail> GetFabricWashTest_Detail(long ID);

        void UpdateFabricWashTestDetail(FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result, string userID);

        void EncodeFabricWash(long ID, string testResult, DateTime? WashDate, string userID);

        DataTable GetWashFailMailContentData(long ID);

        void AmendFabricWash(long ID);

        DataTable GetWashDetailForReport(long ID);
    }
}
