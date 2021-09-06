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

        BaseResult SaveFabricCrkShrkTestCrockingDetail(FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestCrockingDetail(long ID, string userID, out string testResult);

        SendMail_Result SendCrockingFailResultMail(string toAddress, string ccAddress, long ID, bool isTest);

        BaseResult AmendFabricCrkShrkTestCrockingDetail(long ID);

        BaseResult ToExcelFabricCrkShrkTestCrockingDetail(long ID, out string excelFileName, bool isTest);

        BaseResult ToPdfFabricCrkShrkTestCrockingDetail(long ID, out string pdfFileName, bool isTest);

        // Heat
        FabricCrkShrkTestHeat_Result GetFabricCrkShrkTestHeat_Result(long ID);

        BaseResult SaveFabricCrkShrkTestHeatDetail(FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestHeatDetail(long ID, string userID, out string ovenTestResult);
        SendMail_Result SendHeatFailResultMail(string toAddress, string ccAddress, long ID, bool isTest);
        BaseResult AmendFabricCrkShrkTestHeatDetail(long ID);

        BaseResult ToExcelFabricCrkShrkTestHeatDetail(long ID, out string excelFileName, bool isTest);

        BaseResult ToPdfFabricCrkShrkTestHeatDetail(long ID, out string pdfFileName, bool isTest);

        // Wash
        FabricCrkShrkTestWash_Result GetFabricCrkShrkTestWash_Result(long ID);

        BaseResult SaveFabricCrkShrkTestWashDetail(FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result, string userID);

        BaseResult EncodeFabricCrkShrkTestWashDetail(long ID, string userID, out string ovenTestResult);

        BaseResult AmendFabricCrkShrkTestWashDetail(long ID);

        BaseResult ToExcelFabricCrkShrkTestWashDetail(long ID, out string excelFileName, bool isTest);

        BaseResult ToPdfFabricCrkShrkTestWashDetail(long ID, out string pdfFileName, bool isTest);


    }
}
