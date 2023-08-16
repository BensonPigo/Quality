using DatabaseObject;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        SendMail_Result SendCrockingFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID);

        BaseResult AmendFabricCrkShrkTestCrockingDetail(long ID);

        //BaseResult ToExcelFabricCrkShrkTestCrockingDetail(long ID, out string excelFileName, bool isTest);   ISP20220019註解
        BaseResult Crocking_ToExcel(long ID, bool IsPDF, out string excelFileName);
        //BaseResult ToPdfFabricCrkShrkTestCrockingDetail(long ID, out string pdfFileName, bool isTest);   ISP20220019註解

        // Heat
        FabricCrkShrkTestHeat_Result GetFabricCrkShrkTestHeat_Result(long ID);

        BaseResult SaveFabricCrkShrkTestHeatDetail(FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestHeatDetail(long ID, string userID, out string ovenTestResult);
        SendMail_Result SendHeatFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID);
        BaseResult AmendFabricCrkShrkTestHeatDetail(long ID);

        BaseResult ToExcelFabricCrkShrkTestHeatDetail(long ID, out string excelFileName);

        // Wash
        FabricCrkShrkTestWash_Result GetFabricCrkShrkTestWash_Result(long ID);

        BaseResult SaveFabricCrkShrkTestWashDetail(FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestWashDetail(long ID, string userID, out string testResult);

        SendMail_Result SendWashFailResultMail(string toAddress, string ccAddress, long ID, bool isTest, string OrderID);

        BaseResult AmendFabricCrkShrkTestWashDetail(long ID);

        BaseResult ToExcelFabricCrkShrkTestWashDetail(long ID, out string excelFileName);


    }
}
