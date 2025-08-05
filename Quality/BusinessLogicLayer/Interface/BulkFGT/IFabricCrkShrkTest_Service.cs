using DatabaseObject;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace BusinessLogicLayer.Interface
{
    public interface IFabricCrkShrkTest_Service
    {
        FabricCrkShrkTest_Result GetFabricCrkShrkTest_Result(string POID);

        BaseResult SaveFabricCrkShrkTestMain(FabricCrkShrkTest_Result fabricCrkShrkTest_Result);


        // Crocking
        FabricCrkShrkTestCrocking_Result GetFabricCrkShrkTestCrocking_Result(long ID);

        List<string> GetScaleIDs();

        BaseResult SaveFabricCrkShrkTestCrockingDetail(FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestCrockingDetail(long ID, string userID, out string testResult);

        SendMail_Result SendCrockingFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files);

        BaseResult AmendFabricCrkShrkTestCrockingDetail(long ID);

        //BaseResult ToExcelFabricCrkShrkTestCrockingDetail(long ID, out string excelFileName, bool isTest);   ISP20220019註解
        BaseResult ToReport_Crocking(long ID, bool IsPDF, out string excelFileName, string AssignedFineName = "");
        //BaseResult ToPdfFabricCrkShrkTestCrockingDetail(long ID, out string pdfFileName, bool isTest);   ISP20220019註解

        // Heat
        FabricCrkShrkTestHeat_Result GetFabricCrkShrkTestHeat_Result(long ID);

        BaseResult SaveFabricCrkShrkTestHeatDetail(FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestHeatDetail(long ID, string userID, out string ovenTestResult);
        SendMail_Result SendHeatFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files);
        BaseResult AmendFabricCrkShrkTestHeatDetail(long ID);

        BaseResult ToReport_Heat(long ID, bool IsToPDF, out string excelFileName, string AssignedFineName = "");

        // Iron
        FabricCrkShrkTestIron_Result GetFabricCrkShrkTestIron_Result(long ID);

        BaseResult SaveFabricCrkShrkTestIronDetail(FabricCrkShrkTestIron_Result fabricCrkShrkTestIron_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestIronDetail(long ID, string userID, out string ovenTestResult);
        SendMail_Result SendIronFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files);
        BaseResult AmendFabricCrkShrkTestIronDetail(long ID);

        BaseResult ToReport_Iron(long ID, bool IsToPDF, out string excelFileName, string AssignedFineName = "");

        // Wash
        FabricCrkShrkTestWash_Result GetFabricCrkShrkTestWash_Result(long ID);

        BaseResult SaveFabricCrkShrkTestWashDetail(FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestWashDetail(long ID, string userID, out string testResult);

        SendMail_Result SendWashFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files);

        BaseResult AmendFabricCrkShrkTestWashDetail(long ID);

        BaseResult ToReport_Wash(long ID, bool IsToPDF, out string excelFileName, string AssignedFineName = "");

        // Weight
        FabricCrkShrkTestWeight_Result GetFabricCrkShrkTestWeight_Result(long ID);

        BaseResult SaveFabricCrkShrkTestWeightDetail(FabricCrkShrkTestWeight_Result fabricCrkShrkTestWeight_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestWeightDetail(long ID, string userID, out string testResult);

        SendMail_Result SendWeightFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID, string Subject, string Body, List<HttpPostedFileBase> Files);

        BaseResult AmendFabricCrkShrkTestWeightDetail(long ID, string userID);

        BaseResult ToReport_Weight(long ID, bool IsToPDF, out string excelFileName, string AssignedFineName = "");

    }
}
