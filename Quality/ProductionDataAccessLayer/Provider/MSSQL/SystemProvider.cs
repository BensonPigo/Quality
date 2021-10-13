using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using DatabaseObject.ProductionDB;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class SystemProvider : SQLDAL, ISystemProvider
    {
        #region 底層連線
        public SystemProvider(string ConString) : base(ConString) { }
        public SystemProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        /*回傳系統參數檔(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳系統參數檔
        /// </summary>
        /// <param name="Item">系統參數檔成員</param>
        /// <returns>回傳系統參數檔</returns>
        /// <info>Author: Admin; Date: 2021/08/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/12  1.00    Admin        Create
        /// </history>
        public IList<DatabaseObject.ProductionDB.System> Get()
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         Mailserver" + Environment.NewLine);
            SbSql.Append("        ,Sendfrom" + Environment.NewLine);
            SbSql.Append("        ,EmailID" + Environment.NewLine);
            SbSql.Append("        ,EmailPwd" + Environment.NewLine);
            SbSql.Append("        ,PicPath" + Environment.NewLine);
            SbSql.Append("        ,StdTMS" + Environment.NewLine);
            SbSql.Append("        ,ClipPath" + Environment.NewLine);
            SbSql.Append("        ,FtpIP" + Environment.NewLine);
            SbSql.Append("        ,FtpID" + Environment.NewLine);
            SbSql.Append("        ,FtpPwd" + Environment.NewLine);
            SbSql.Append("        ,SewLock" + Environment.NewLine);
            SbSql.Append("        ,SampleRate" + Environment.NewLine);
            SbSql.Append("        ,PullLock" + Environment.NewLine);
            SbSql.Append("        ,RgCode" + Environment.NewLine);
            SbSql.Append("        ,ImportDataPath" + Environment.NewLine);
            SbSql.Append("        ,ImportDataFileName" + Environment.NewLine);
            SbSql.Append("        ,ExportDataPath" + Environment.NewLine);
            SbSql.Append("        ,CurrencyID" + Environment.NewLine);
            SbSql.Append("        ,USDRate" + Environment.NewLine);
            SbSql.Append("        ,POApproveName" + Environment.NewLine);
            SbSql.Append("        ,POApproveDay" + Environment.NewLine);
            SbSql.Append("        ,CutDay" + Environment.NewLine);
            SbSql.Append("        ,AccountKeyword" + Environment.NewLine);
            SbSql.Append("        ,ReadyDay" + Environment.NewLine);
            SbSql.Append("        ,VNMultiple" + Environment.NewLine);
            SbSql.Append("        ,MtlLeadTime" + Environment.NewLine);
            SbSql.Append("        ,ExchangeID" + Environment.NewLine);
            SbSql.Append("        ,RFIDServerName" + Environment.NewLine);
            SbSql.Append("        ,RFIDDatabaseName" + Environment.NewLine);
            SbSql.Append("        ,RFIDLoginId" + Environment.NewLine);
            SbSql.Append("        ,RFIDLoginPwd" + Environment.NewLine);
            SbSql.Append("        ,RFIDTable" + Environment.NewLine);
            SbSql.Append("        ,ProphetSingleSizeDeduct" + Environment.NewLine);
            SbSql.Append("        ,PrintingSuppID" + Environment.NewLine);
            SbSql.Append("        ,QCMachineDelayTime" + Environment.NewLine);
            SbSql.Append("        ,APSLoginId" + Environment.NewLine);
            SbSql.Append("        ,APSLoginPwd" + Environment.NewLine);
            SbSql.Append("        ,SQLServerName" + Environment.NewLine);
            SbSql.Append("        ,APSDatabaseName" + Environment.NewLine);
            SbSql.Append("        ,RFIDMiddlewareInRFIDServer" + Environment.NewLine);
            SbSql.Append("        ,UseAutoScanPack" + Environment.NewLine);
            SbSql.Append("        ,MtlAutoLock" + Environment.NewLine);
            SbSql.Append("        ,InspAutoLockAcc" + Environment.NewLine);
            SbSql.Append("        ,ShippingMarkPath" + Environment.NewLine);
            SbSql.Append("        ,StyleSketch" + Environment.NewLine);
            SbSql.Append("        ,ARKServerName" + Environment.NewLine);
            SbSql.Append("        ,ARKDatabaseName" + Environment.NewLine);
            SbSql.Append("        ,ARKLoginId" + Environment.NewLine);
            SbSql.Append("        ,ARKLoginPwd" + Environment.NewLine);
            SbSql.Append("        ,MarkerInputPath" + Environment.NewLine);
            SbSql.Append("        ,MarkerOutputPath" + Environment.NewLine);
            SbSql.Append("        ,ReplacementReport" + Environment.NewLine);
            SbSql.Append("        ,CuttingP10mustCutRef" + Environment.NewLine);
            SbSql.Append("        ,Automation" + Environment.NewLine);
            SbSql.Append("        ,AutomationAutoRunTime" + Environment.NewLine);
            SbSql.Append("        ,CanReviseDailyLockData" + Environment.NewLine);
            SbSql.Append("        ,AutoGenerateByTone" + Environment.NewLine);
            SbSql.Append("        ,MiscPOApproveName" + Environment.NewLine);
            SbSql.Append("        ,MiscPOApproveDay" + Environment.NewLine);
            SbSql.Append("        ,QMSAutoAdjustMtl" + Environment.NewLine);
            SbSql.Append("        ,ShippingMarkTemplatePath" + Environment.NewLine);
            SbSql.Append("        ,WIP_FollowCutOutput" + Environment.NewLine);
            SbSql.Append("        ,NoRestrictOrdersDelivery" + Environment.NewLine);
            SbSql.Append("        ,WIP_ByShell" + Environment.NewLine);
            SbSql.Append("        ,RFCardEraseBeforePrinting" + Environment.NewLine);
            SbSql.Append("        ,SewlineAvgCPU" + Environment.NewLine);
            SbSql.Append("        ,SmallLogoCM" + Environment.NewLine);
            SbSql.Append("        ,CheckRFIDCardDuplicateByWebservice" + Environment.NewLine);
            SbSql.Append("        ,IsCombineSubProcess" + Environment.NewLine);
            SbSql.Append("        ,IsNoneShellNoCreateAllParts" + Environment.NewLine);
            SbSql.Append("        ,Region" + Environment.NewLine);
            SbSql.Append("        ,DQSQtyPCT" + Environment.NewLine);
            SbSql.Append("        ,FinalInspection_CTNMoistureStandard" + Environment.NewLine);
            SbSql.Append("FROM [System]" + Environment.NewLine);



            return ExecuteList<DatabaseObject.ProductionDB.System>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        /*建立系統參數檔(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立系統參數檔
        /// </summary>
        /// <param name="Item">系統參數檔成員</param>
        /// <returns>回傳異動筆數</returns>
        /// <info>Author: Admin; Date: 2021/08/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/12  1.00    Admin        Create
        /// </history>
        public int Create(DatabaseObject.ProductionDB.System Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [System]" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         Mailserver" + Environment.NewLine);
            SbSql.Append("        ,Sendfrom" + Environment.NewLine);
            SbSql.Append("        ,EmailID" + Environment.NewLine);
            SbSql.Append("        ,EmailPwd" + Environment.NewLine);
            SbSql.Append("        ,PicPath" + Environment.NewLine);
            SbSql.Append("        ,StdTMS" + Environment.NewLine);
            SbSql.Append("        ,ClipPath" + Environment.NewLine);
            SbSql.Append("        ,FtpIP" + Environment.NewLine);
            SbSql.Append("        ,FtpID" + Environment.NewLine);
            SbSql.Append("        ,FtpPwd" + Environment.NewLine);
            SbSql.Append("        ,SewLock" + Environment.NewLine);
            SbSql.Append("        ,SampleRate" + Environment.NewLine);
            SbSql.Append("        ,PullLock" + Environment.NewLine);
            SbSql.Append("        ,RgCode" + Environment.NewLine);
            SbSql.Append("        ,ImportDataPath" + Environment.NewLine);
            SbSql.Append("        ,ImportDataFileName" + Environment.NewLine);
            SbSql.Append("        ,ExportDataPath" + Environment.NewLine);
            SbSql.Append("        ,CurrencyID" + Environment.NewLine);
            SbSql.Append("        ,USDRate" + Environment.NewLine);
            SbSql.Append("        ,POApproveName" + Environment.NewLine);
            SbSql.Append("        ,POApproveDay" + Environment.NewLine);
            SbSql.Append("        ,CutDay" + Environment.NewLine);
            SbSql.Append("        ,AccountKeyword" + Environment.NewLine);
            SbSql.Append("        ,ReadyDay" + Environment.NewLine);
            SbSql.Append("        ,VNMultiple" + Environment.NewLine);
            SbSql.Append("        ,MtlLeadTime" + Environment.NewLine);
            SbSql.Append("        ,ExchangeID" + Environment.NewLine);
            SbSql.Append("        ,RFIDServerName" + Environment.NewLine);
            SbSql.Append("        ,RFIDDatabaseName" + Environment.NewLine);
            SbSql.Append("        ,RFIDLoginId" + Environment.NewLine);
            SbSql.Append("        ,RFIDLoginPwd" + Environment.NewLine);
            SbSql.Append("        ,RFIDTable" + Environment.NewLine);
            SbSql.Append("        ,ProphetSingleSizeDeduct" + Environment.NewLine);
            SbSql.Append("        ,PrintingSuppID" + Environment.NewLine);
            SbSql.Append("        ,QCMachineDelayTime" + Environment.NewLine);
            SbSql.Append("        ,APSLoginId" + Environment.NewLine);
            SbSql.Append("        ,APSLoginPwd" + Environment.NewLine);
            SbSql.Append("        ,SQLServerName" + Environment.NewLine);
            SbSql.Append("        ,APSDatabaseName" + Environment.NewLine);
            SbSql.Append("        ,RFIDMiddlewareInRFIDServer" + Environment.NewLine);
            SbSql.Append("        ,UseAutoScanPack" + Environment.NewLine);
            SbSql.Append("        ,MtlAutoLock" + Environment.NewLine);
            SbSql.Append("        ,InspAutoLockAcc" + Environment.NewLine);
            SbSql.Append("        ,ShippingMarkPath" + Environment.NewLine);
            SbSql.Append("        ,StyleSketch" + Environment.NewLine);
            SbSql.Append("        ,ARKServerName" + Environment.NewLine);
            SbSql.Append("        ,ARKDatabaseName" + Environment.NewLine);
            SbSql.Append("        ,ARKLoginId" + Environment.NewLine);
            SbSql.Append("        ,ARKLoginPwd" + Environment.NewLine);
            SbSql.Append("        ,MarkerInputPath" + Environment.NewLine);
            SbSql.Append("        ,MarkerOutputPath" + Environment.NewLine);
            SbSql.Append("        ,ReplacementReport" + Environment.NewLine);
            SbSql.Append("        ,CuttingP10mustCutRef" + Environment.NewLine);
            SbSql.Append("        ,Automation" + Environment.NewLine);
            SbSql.Append("        ,AutomationAutoRunTime" + Environment.NewLine);
            SbSql.Append("        ,CanReviseDailyLockData" + Environment.NewLine);
            SbSql.Append("        ,AutoGenerateByTone" + Environment.NewLine);
            SbSql.Append("        ,MiscPOApproveName" + Environment.NewLine);
            SbSql.Append("        ,MiscPOApproveDay" + Environment.NewLine);
            SbSql.Append("        ,QMSAutoAdjustMtl" + Environment.NewLine);
            SbSql.Append("        ,ShippingMarkTemplatePath" + Environment.NewLine);
            SbSql.Append("        ,WIP_FollowCutOutput" + Environment.NewLine);
            SbSql.Append("        ,NoRestrictOrdersDelivery" + Environment.NewLine);
            SbSql.Append("        ,WIP_ByShell" + Environment.NewLine);
            SbSql.Append("        ,RFCardEraseBeforePrinting" + Environment.NewLine);
            SbSql.Append("        ,SewlineAvgCPU" + Environment.NewLine);
            SbSql.Append("        ,SmallLogoCM" + Environment.NewLine);
            SbSql.Append("        ,CheckRFIDCardDuplicateByWebservice" + Environment.NewLine);
            SbSql.Append("        ,IsCombineSubProcess" + Environment.NewLine);
            SbSql.Append("        ,IsNoneShellNoCreateAllParts" + Environment.NewLine);
            SbSql.Append("        ,Region" + Environment.NewLine);
            SbSql.Append("        ,DQSQtyPCT" + Environment.NewLine);
            SbSql.Append(")" + Environment.NewLine);
            SbSql.Append("VALUES" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @Mailserver"); objParameter.Add("@Mailserver", DbType.String, Item.Mailserver);
            SbSql.Append("        ,@Sendfrom"); objParameter.Add("@Sendfrom", DbType.String, Item.Sendfrom);
            SbSql.Append("        ,@EmailID"); objParameter.Add("@EmailID", DbType.String, Item.EmailID);
            SbSql.Append("        ,@EmailPwd"); objParameter.Add("@EmailPwd", DbType.String, Item.EmailPwd);
            SbSql.Append("        ,@PicPath"); objParameter.Add("@PicPath", DbType.String, Item.PicPath);
            SbSql.Append("        ,@StdTMS"); objParameter.Add("@StdTMS", DbType.Int32, Item.StdTMS);
            SbSql.Append("        ,@ClipPath"); objParameter.Add("@ClipPath", DbType.String, Item.ClipPath);
            SbSql.Append("        ,@FtpIP"); objParameter.Add("@FtpIP", DbType.String, Item.FtpIP);
            SbSql.Append("        ,@FtpID"); objParameter.Add("@FtpID", DbType.String, Item.FtpID);
            SbSql.Append("        ,@FtpPwd"); objParameter.Add("@FtpPwd", DbType.String, Item.FtpPwd);
            SbSql.Append("        ,@SewLock"); objParameter.Add("@SewLock", DbType.String, Item.SewLock);
            SbSql.Append("        ,@SampleRate"); objParameter.Add("@SampleRate", DbType.String, Item.SampleRate);
            SbSql.Append("        ,@PullLock"); objParameter.Add("@PullLock", DbType.String, Item.PullLock);
            SbSql.Append("        ,@RgCode"); objParameter.Add("@RgCode", DbType.String, Item.RgCode);
            SbSql.Append("        ,@ImportDataPath"); objParameter.Add("@ImportDataPath", DbType.String, Item.ImportDataPath);
            SbSql.Append("        ,@ImportDataFileName"); objParameter.Add("@ImportDataFileName", DbType.String, Item.ImportDataFileName);
            SbSql.Append("        ,@ExportDataPath"); objParameter.Add("@ExportDataPath", DbType.String, Item.ExportDataPath);
            SbSql.Append("        ,@CurrencyID"); objParameter.Add("@CurrencyID", DbType.String, Item.CurrencyID);
            SbSql.Append("        ,@USDRate"); objParameter.Add("@USDRate", DbType.String, Item.USDRate);
            SbSql.Append("        ,@POApproveName"); objParameter.Add("@POApproveName", DbType.String, Item.POApproveName);
            SbSql.Append("        ,@POApproveDay"); objParameter.Add("@POApproveDay", DbType.String, Item.POApproveDay);
            SbSql.Append("        ,@CutDay"); objParameter.Add("@CutDay", DbType.String, Item.CutDay);
            SbSql.Append("        ,@AccountKeyword"); objParameter.Add("@AccountKeyword", DbType.String, Item.AccountKeyword);
            SbSql.Append("        ,@ReadyDay"); objParameter.Add("@ReadyDay", DbType.String, Item.ReadyDay);
            SbSql.Append("        ,@VNMultiple"); objParameter.Add("@VNMultiple", DbType.String, Item.VNMultiple);
            SbSql.Append("        ,@MtlLeadTime"); objParameter.Add("@MtlLeadTime", DbType.String, Item.MtlLeadTime);
            SbSql.Append("        ,@ExchangeID"); objParameter.Add("@ExchangeID", DbType.String, Item.ExchangeID);
            SbSql.Append("        ,@RFIDServerName"); objParameter.Add("@RFIDServerName", DbType.String, Item.RFIDServerName);
            SbSql.Append("        ,@RFIDDatabaseName"); objParameter.Add("@RFIDDatabaseName", DbType.String, Item.RFIDDatabaseName);
            SbSql.Append("        ,@RFIDLoginId"); objParameter.Add("@RFIDLoginId", DbType.String, Item.RFIDLoginId);
            SbSql.Append("        ,@RFIDLoginPwd"); objParameter.Add("@RFIDLoginPwd", DbType.String, Item.RFIDLoginPwd);
            SbSql.Append("        ,@RFIDTable"); objParameter.Add("@RFIDTable", DbType.String, Item.RFIDTable);
            SbSql.Append("        ,@ProphetSingleSizeDeduct"); objParameter.Add("@ProphetSingleSizeDeduct", DbType.String, Item.ProphetSingleSizeDeduct);
            SbSql.Append("        ,@PrintingSuppID"); objParameter.Add("@PrintingSuppID", DbType.String, Item.PrintingSuppID);
            SbSql.Append("        ,@QCMachineDelayTime"); objParameter.Add("@QCMachineDelayTime", DbType.String, Item.QCMachineDelayTime);
            SbSql.Append("        ,@APSLoginId"); objParameter.Add("@APSLoginId", DbType.String, Item.APSLoginId);
            SbSql.Append("        ,@APSLoginPwd"); objParameter.Add("@APSLoginPwd", DbType.String, Item.APSLoginPwd);
            SbSql.Append("        ,@SQLServerName"); objParameter.Add("@SQLServerName", DbType.String, Item.SQLServerName);
            SbSql.Append("        ,@APSDatabaseName"); objParameter.Add("@APSDatabaseName", DbType.String, Item.APSDatabaseName);
            SbSql.Append("        ,@RFIDMiddlewareInRFIDServer"); objParameter.Add("@RFIDMiddlewareInRFIDServer", DbType.String, Item.RFIDMiddlewareInRFIDServer);
            SbSql.Append("        ,@UseAutoScanPack"); objParameter.Add("@UseAutoScanPack", DbType.String, Item.UseAutoScanPack);
            SbSql.Append("        ,@MtlAutoLock"); objParameter.Add("@MtlAutoLock", DbType.String, Item.MtlAutoLock);
            SbSql.Append("        ,@InspAutoLockAcc"); objParameter.Add("@InspAutoLockAcc", DbType.String, Item.InspAutoLockAcc);
            SbSql.Append("        ,@ShippingMarkPath"); objParameter.Add("@ShippingMarkPath", DbType.String, Item.ShippingMarkPath);
            SbSql.Append("        ,@StyleSketch"); objParameter.Add("@StyleSketch", DbType.String, Item.StyleSketch);
            SbSql.Append("        ,@ARKServerName"); objParameter.Add("@ARKServerName", DbType.String, Item.ARKServerName);
            SbSql.Append("        ,@ARKDatabaseName"); objParameter.Add("@ARKDatabaseName", DbType.String, Item.ARKDatabaseName);
            SbSql.Append("        ,@ARKLoginId"); objParameter.Add("@ARKLoginId", DbType.String, Item.ARKLoginId);
            SbSql.Append("        ,@ARKLoginPwd"); objParameter.Add("@ARKLoginPwd", DbType.String, Item.ARKLoginPwd);
            SbSql.Append("        ,@MarkerInputPath"); objParameter.Add("@MarkerInputPath", DbType.String, Item.MarkerInputPath);
            SbSql.Append("        ,@MarkerOutputPath"); objParameter.Add("@MarkerOutputPath", DbType.String, Item.MarkerOutputPath);
            SbSql.Append("        ,@ReplacementReport"); objParameter.Add("@ReplacementReport", DbType.String, Item.ReplacementReport);
            SbSql.Append("        ,@CuttingP10mustCutRef"); objParameter.Add("@CuttingP10mustCutRef", DbType.String, Item.CuttingP10mustCutRef);
            SbSql.Append("        ,@Automation"); objParameter.Add("@Automation", DbType.String, Item.Automation);
            SbSql.Append("        ,@AutomationAutoRunTime"); objParameter.Add("@AutomationAutoRunTime", DbType.String, Item.AutomationAutoRunTime);
            SbSql.Append("        ,@CanReviseDailyLockData"); objParameter.Add("@CanReviseDailyLockData", DbType.String, Item.CanReviseDailyLockData);
            SbSql.Append("        ,@AutoGenerateByTone"); objParameter.Add("@AutoGenerateByTone", DbType.String, Item.AutoGenerateByTone);
            SbSql.Append("        ,@MiscPOApproveName"); objParameter.Add("@MiscPOApproveName", DbType.String, Item.MiscPOApproveName);
            SbSql.Append("        ,@MiscPOApproveDay"); objParameter.Add("@MiscPOApproveDay", DbType.String, Item.MiscPOApproveDay);
            SbSql.Append("        ,@QMSAutoAdjustMtl"); objParameter.Add("@QMSAutoAdjustMtl", DbType.String, Item.QMSAutoAdjustMtl);
            SbSql.Append("        ,@ShippingMarkTemplatePath"); objParameter.Add("@ShippingMarkTemplatePath", DbType.String, Item.ShippingMarkTemplatePath);
            SbSql.Append("        ,@WIP_FollowCutOutput"); objParameter.Add("@WIP_FollowCutOutput", DbType.String, Item.WIP_FollowCutOutput);
            SbSql.Append("        ,@NoRestrictOrdersDelivery"); objParameter.Add("@NoRestrictOrdersDelivery", DbType.String, Item.NoRestrictOrdersDelivery);
            SbSql.Append("        ,@WIP_ByShell"); objParameter.Add("@WIP_ByShell", DbType.String, Item.WIP_ByShell);
            SbSql.Append("        ,@RFCardEraseBeforePrinting"); objParameter.Add("@RFCardEraseBeforePrinting", DbType.String, Item.RFCardEraseBeforePrinting);
            SbSql.Append("        ,@SewlineAvgCPU"); objParameter.Add("@SewlineAvgCPU", DbType.Int32, Item.SewlineAvgCPU);
            SbSql.Append("        ,@SmallLogoCM"); objParameter.Add("@SmallLogoCM", DbType.Decimal, Item.SmallLogoCM);
            SbSql.Append("        ,@CheckRFIDCardDuplicateByWebservice"); objParameter.Add("@CheckRFIDCardDuplicateByWebservice", DbType.String, Item.CheckRFIDCardDuplicateByWebservice);
            SbSql.Append("        ,@IsCombineSubProcess"); objParameter.Add("@IsCombineSubProcess", DbType.String, Item.IsCombineSubProcess);
            SbSql.Append("        ,@IsNoneShellNoCreateAllParts"); objParameter.Add("@IsNoneShellNoCreateAllParts", DbType.String, Item.IsNoneShellNoCreateAllParts);
            SbSql.Append("        ,@Region"); objParameter.Add("@Region", DbType.String, Item.Region);
            SbSql.Append("        ,@DQSQtyPCT"); objParameter.Add("@DQSQtyPCT", DbType.String, Item.DQSQtyPCT);
            SbSql.Append(")" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        /*更新系統參數檔(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新系統參數檔
        /// </summary>
        /// <param name="Item">系統參數檔成員</param>
        /// <returns>回傳異動筆數</returns>
        /// <info>Author: Admin; Date: 2021/08/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/12  1.00    Admin        Create
        /// </history>
        public int Update(DatabaseObject.ProductionDB.System Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [System]" + Environment.NewLine);
            SbSql.Append("SET" + Environment.NewLine);
            if (Item.Mailserver != null) { SbSql.Append("Mailserver=@Mailserver" + Environment.NewLine); objParameter.Add("@Mailserver", DbType.String, Item.Mailserver); }
            if (Item.Sendfrom != null) { SbSql.Append(",Sendfrom=@Sendfrom" + Environment.NewLine); objParameter.Add("@Sendfrom", DbType.String, Item.Sendfrom); }
            if (Item.EmailID != null) { SbSql.Append(",EmailID=@EmailID" + Environment.NewLine); objParameter.Add("@EmailID", DbType.String, Item.EmailID); }
            if (Item.EmailPwd != null) { SbSql.Append(",EmailPwd=@EmailPwd" + Environment.NewLine); objParameter.Add("@EmailPwd", DbType.String, Item.EmailPwd); }
            if (Item.PicPath != null) { SbSql.Append(",PicPath=@PicPath" + Environment.NewLine); objParameter.Add("@PicPath", DbType.String, Item.PicPath); }
            if (Item.StdTMS != null) { SbSql.Append(",StdTMS=@StdTMS" + Environment.NewLine); objParameter.Add("@StdTMS", DbType.Int32, Item.StdTMS); }
            if (Item.ClipPath != null) { SbSql.Append(",ClipPath=@ClipPath" + Environment.NewLine); objParameter.Add("@ClipPath", DbType.String, Item.ClipPath); }
            if (Item.FtpIP != null) { SbSql.Append(",FtpIP=@FtpIP" + Environment.NewLine); objParameter.Add("@FtpIP", DbType.String, Item.FtpIP); }
            if (Item.FtpID != null) { SbSql.Append(",FtpID=@FtpID" + Environment.NewLine); objParameter.Add("@FtpID", DbType.String, Item.FtpID); }
            if (Item.FtpPwd != null) { SbSql.Append(",FtpPwd=@FtpPwd" + Environment.NewLine); objParameter.Add("@FtpPwd", DbType.String, Item.FtpPwd); }
            if (Item.SewLock != null) { SbSql.Append(",SewLock=@SewLock" + Environment.NewLine); objParameter.Add("@SewLock", DbType.String, Item.SewLock); }
            if (Item.SampleRate != null) { SbSql.Append(",SampleRate=@SampleRate" + Environment.NewLine); objParameter.Add("@SampleRate", DbType.String, Item.SampleRate); }
            if (Item.PullLock != null) { SbSql.Append(",PullLock=@PullLock" + Environment.NewLine); objParameter.Add("@PullLock", DbType.String, Item.PullLock); }
            if (Item.RgCode != null) { SbSql.Append(",RgCode=@RgCode" + Environment.NewLine); objParameter.Add("@RgCode", DbType.String, Item.RgCode); }
            if (Item.ImportDataPath != null) { SbSql.Append(",ImportDataPath=@ImportDataPath" + Environment.NewLine); objParameter.Add("@ImportDataPath", DbType.String, Item.ImportDataPath); }
            if (Item.ImportDataFileName != null) { SbSql.Append(",ImportDataFileName=@ImportDataFileName" + Environment.NewLine); objParameter.Add("@ImportDataFileName", DbType.String, Item.ImportDataFileName); }
            if (Item.ExportDataPath != null) { SbSql.Append(",ExportDataPath=@ExportDataPath" + Environment.NewLine); objParameter.Add("@ExportDataPath", DbType.String, Item.ExportDataPath); }
            if (Item.CurrencyID != null) { SbSql.Append(",CurrencyID=@CurrencyID" + Environment.NewLine); objParameter.Add("@CurrencyID", DbType.String, Item.CurrencyID); }
            if (Item.USDRate != null) { SbSql.Append(",USDRate=@USDRate" + Environment.NewLine); objParameter.Add("@USDRate", DbType.String, Item.USDRate); }
            if (Item.POApproveName != null) { SbSql.Append(",POApproveName=@POApproveName" + Environment.NewLine); objParameter.Add("@POApproveName", DbType.String, Item.POApproveName); }
            if (Item.POApproveDay != null) { SbSql.Append(",POApproveDay=@POApproveDay" + Environment.NewLine); objParameter.Add("@POApproveDay", DbType.String, Item.POApproveDay); }
            if (Item.CutDay != null) { SbSql.Append(",CutDay=@CutDay" + Environment.NewLine); objParameter.Add("@CutDay", DbType.String, Item.CutDay); }
            if (Item.AccountKeyword != null) { SbSql.Append(",AccountKeyword=@AccountKeyword" + Environment.NewLine); objParameter.Add("@AccountKeyword", DbType.String, Item.AccountKeyword); }
            if (Item.ReadyDay != null) { SbSql.Append(",ReadyDay=@ReadyDay" + Environment.NewLine); objParameter.Add("@ReadyDay", DbType.String, Item.ReadyDay); }
            if (Item.VNMultiple != null) { SbSql.Append(",VNMultiple=@VNMultiple" + Environment.NewLine); objParameter.Add("@VNMultiple", DbType.String, Item.VNMultiple); }
            if (Item.MtlLeadTime != null) { SbSql.Append(",MtlLeadTime=@MtlLeadTime" + Environment.NewLine); objParameter.Add("@MtlLeadTime", DbType.String, Item.MtlLeadTime); }
            if (Item.ExchangeID != null) { SbSql.Append(",ExchangeID=@ExchangeID" + Environment.NewLine); objParameter.Add("@ExchangeID", DbType.String, Item.ExchangeID); }
            if (Item.RFIDServerName != null) { SbSql.Append(",RFIDServerName=@RFIDServerName" + Environment.NewLine); objParameter.Add("@RFIDServerName", DbType.String, Item.RFIDServerName); }
            if (Item.RFIDDatabaseName != null) { SbSql.Append(",RFIDDatabaseName=@RFIDDatabaseName" + Environment.NewLine); objParameter.Add("@RFIDDatabaseName", DbType.String, Item.RFIDDatabaseName); }
            if (Item.RFIDLoginId != null) { SbSql.Append(",RFIDLoginId=@RFIDLoginId" + Environment.NewLine); objParameter.Add("@RFIDLoginId", DbType.String, Item.RFIDLoginId); }
            if (Item.RFIDLoginPwd != null) { SbSql.Append(",RFIDLoginPwd=@RFIDLoginPwd" + Environment.NewLine); objParameter.Add("@RFIDLoginPwd", DbType.String, Item.RFIDLoginPwd); }
            if (Item.RFIDTable != null) { SbSql.Append(",RFIDTable=@RFIDTable" + Environment.NewLine); objParameter.Add("@RFIDTable", DbType.String, Item.RFIDTable); }
            if (Item.ProphetSingleSizeDeduct != null) { SbSql.Append(",ProphetSingleSizeDeduct=@ProphetSingleSizeDeduct" + Environment.NewLine); objParameter.Add("@ProphetSingleSizeDeduct", DbType.String, Item.ProphetSingleSizeDeduct); }
            if (Item.PrintingSuppID != null) { SbSql.Append(",PrintingSuppID=@PrintingSuppID" + Environment.NewLine); objParameter.Add("@PrintingSuppID", DbType.String, Item.PrintingSuppID); }
            if (Item.QCMachineDelayTime != null) { SbSql.Append(",QCMachineDelayTime=@QCMachineDelayTime" + Environment.NewLine); objParameter.Add("@QCMachineDelayTime", DbType.String, Item.QCMachineDelayTime); }
            if (Item.APSLoginId != null) { SbSql.Append(",APSLoginId=@APSLoginId" + Environment.NewLine); objParameter.Add("@APSLoginId", DbType.String, Item.APSLoginId); }
            if (Item.APSLoginPwd != null) { SbSql.Append(",APSLoginPwd=@APSLoginPwd" + Environment.NewLine); objParameter.Add("@APSLoginPwd", DbType.String, Item.APSLoginPwd); }
            if (Item.SQLServerName != null) { SbSql.Append(",SQLServerName=@SQLServerName" + Environment.NewLine); objParameter.Add("@SQLServerName", DbType.String, Item.SQLServerName); }
            if (Item.APSDatabaseName != null) { SbSql.Append(",APSDatabaseName=@APSDatabaseName" + Environment.NewLine); objParameter.Add("@APSDatabaseName", DbType.String, Item.APSDatabaseName); }
            if (Item.RFIDMiddlewareInRFIDServer != null) { SbSql.Append(",RFIDMiddlewareInRFIDServer=@RFIDMiddlewareInRFIDServer" + Environment.NewLine); objParameter.Add("@RFIDMiddlewareInRFIDServer", DbType.String, Item.RFIDMiddlewareInRFIDServer); }
            if (Item.UseAutoScanPack != null) { SbSql.Append(",UseAutoScanPack=@UseAutoScanPack" + Environment.NewLine); objParameter.Add("@UseAutoScanPack", DbType.String, Item.UseAutoScanPack); }
            if (Item.MtlAutoLock != null) { SbSql.Append(",MtlAutoLock=@MtlAutoLock" + Environment.NewLine); objParameter.Add("@MtlAutoLock", DbType.String, Item.MtlAutoLock); }
            if (Item.InspAutoLockAcc != null) { SbSql.Append(",InspAutoLockAcc=@InspAutoLockAcc" + Environment.NewLine); objParameter.Add("@InspAutoLockAcc", DbType.String, Item.InspAutoLockAcc); }
            if (Item.ShippingMarkPath != null) { SbSql.Append(",ShippingMarkPath=@ShippingMarkPath" + Environment.NewLine); objParameter.Add("@ShippingMarkPath", DbType.String, Item.ShippingMarkPath); }
            if (Item.StyleSketch != null) { SbSql.Append(",StyleSketch=@StyleSketch" + Environment.NewLine); objParameter.Add("@StyleSketch", DbType.String, Item.StyleSketch); }
            if (Item.ARKServerName != null) { SbSql.Append(",ARKServerName=@ARKServerName" + Environment.NewLine); objParameter.Add("@ARKServerName", DbType.String, Item.ARKServerName); }
            if (Item.ARKDatabaseName != null) { SbSql.Append(",ARKDatabaseName=@ARKDatabaseName" + Environment.NewLine); objParameter.Add("@ARKDatabaseName", DbType.String, Item.ARKDatabaseName); }
            if (Item.ARKLoginId != null) { SbSql.Append(",ARKLoginId=@ARKLoginId" + Environment.NewLine); objParameter.Add("@ARKLoginId", DbType.String, Item.ARKLoginId); }
            if (Item.ARKLoginPwd != null) { SbSql.Append(",ARKLoginPwd=@ARKLoginPwd" + Environment.NewLine); objParameter.Add("@ARKLoginPwd", DbType.String, Item.ARKLoginPwd); }
            if (Item.MarkerInputPath != null) { SbSql.Append(",MarkerInputPath=@MarkerInputPath" + Environment.NewLine); objParameter.Add("@MarkerInputPath", DbType.String, Item.MarkerInputPath); }
            if (Item.MarkerOutputPath != null) { SbSql.Append(",MarkerOutputPath=@MarkerOutputPath" + Environment.NewLine); objParameter.Add("@MarkerOutputPath", DbType.String, Item.MarkerOutputPath); }
            if (Item.ReplacementReport != null) { SbSql.Append(",ReplacementReport=@ReplacementReport" + Environment.NewLine); objParameter.Add("@ReplacementReport", DbType.String, Item.ReplacementReport); }
            if (Item.CuttingP10mustCutRef != null) { SbSql.Append(",CuttingP10mustCutRef=@CuttingP10mustCutRef" + Environment.NewLine); objParameter.Add("@CuttingP10mustCutRef", DbType.String, Item.CuttingP10mustCutRef); }
            if (Item.Automation != null) { SbSql.Append(",Automation=@Automation" + Environment.NewLine); objParameter.Add("@Automation", DbType.String, Item.Automation); }
            if (Item.AutomationAutoRunTime != null) { SbSql.Append(",AutomationAutoRunTime=@AutomationAutoRunTime" + Environment.NewLine); objParameter.Add("@AutomationAutoRunTime", DbType.String, Item.AutomationAutoRunTime); }
            if (Item.CanReviseDailyLockData != null) { SbSql.Append(",CanReviseDailyLockData=@CanReviseDailyLockData" + Environment.NewLine); objParameter.Add("@CanReviseDailyLockData", DbType.String, Item.CanReviseDailyLockData); }
            if (Item.AutoGenerateByTone != null) { SbSql.Append(",AutoGenerateByTone=@AutoGenerateByTone" + Environment.NewLine); objParameter.Add("@AutoGenerateByTone", DbType.String, Item.AutoGenerateByTone); }
            if (Item.MiscPOApproveName != null) { SbSql.Append(",MiscPOApproveName=@MiscPOApproveName" + Environment.NewLine); objParameter.Add("@MiscPOApproveName", DbType.String, Item.MiscPOApproveName); }
            if (Item.MiscPOApproveDay != null) { SbSql.Append(",MiscPOApproveDay=@MiscPOApproveDay" + Environment.NewLine); objParameter.Add("@MiscPOApproveDay", DbType.String, Item.MiscPOApproveDay); }
            if (Item.QMSAutoAdjustMtl != null) { SbSql.Append(",QMSAutoAdjustMtl=@QMSAutoAdjustMtl" + Environment.NewLine); objParameter.Add("@QMSAutoAdjustMtl", DbType.String, Item.QMSAutoAdjustMtl); }
            if (Item.ShippingMarkTemplatePath != null) { SbSql.Append(",ShippingMarkTemplatePath=@ShippingMarkTemplatePath" + Environment.NewLine); objParameter.Add("@ShippingMarkTemplatePath", DbType.String, Item.ShippingMarkTemplatePath); }
            if (Item.WIP_FollowCutOutput != null) { SbSql.Append(",WIP_FollowCutOutput=@WIP_FollowCutOutput" + Environment.NewLine); objParameter.Add("@WIP_FollowCutOutput", DbType.String, Item.WIP_FollowCutOutput); }
            if (Item.NoRestrictOrdersDelivery != null) { SbSql.Append(",NoRestrictOrdersDelivery=@NoRestrictOrdersDelivery" + Environment.NewLine); objParameter.Add("@NoRestrictOrdersDelivery", DbType.String, Item.NoRestrictOrdersDelivery); }
            if (Item.WIP_ByShell != null) { SbSql.Append(",WIP_ByShell=@WIP_ByShell" + Environment.NewLine); objParameter.Add("@WIP_ByShell", DbType.String, Item.WIP_ByShell); }
            if (Item.RFCardEraseBeforePrinting != null) { SbSql.Append(",RFCardEraseBeforePrinting=@RFCardEraseBeforePrinting" + Environment.NewLine); objParameter.Add("@RFCardEraseBeforePrinting", DbType.String, Item.RFCardEraseBeforePrinting); }
            if (Item.SewlineAvgCPU != null) { SbSql.Append(",SewlineAvgCPU=@SewlineAvgCPU" + Environment.NewLine); objParameter.Add("@SewlineAvgCPU", DbType.Int32, Item.SewlineAvgCPU); }
            if (Item.SmallLogoCM != null) { SbSql.Append(",SmallLogoCM=@SmallLogoCM" + Environment.NewLine); objParameter.Add("@SmallLogoCM", DbType.Decimal, Item.SmallLogoCM); }
            if (Item.CheckRFIDCardDuplicateByWebservice != null) { SbSql.Append(",CheckRFIDCardDuplicateByWebservice=@CheckRFIDCardDuplicateByWebservice" + Environment.NewLine); objParameter.Add("@CheckRFIDCardDuplicateByWebservice", DbType.String, Item.CheckRFIDCardDuplicateByWebservice); }
            if (Item.IsCombineSubProcess != null) { SbSql.Append(",IsCombineSubProcess=@IsCombineSubProcess" + Environment.NewLine); objParameter.Add("@IsCombineSubProcess", DbType.String, Item.IsCombineSubProcess); }
            if (Item.IsNoneShellNoCreateAllParts != null) { SbSql.Append(",IsNoneShellNoCreateAllParts=@IsNoneShellNoCreateAllParts" + Environment.NewLine); objParameter.Add("@IsNoneShellNoCreateAllParts", DbType.String, Item.IsNoneShellNoCreateAllParts); }
            if (Item.Region != null) { SbSql.Append(",Region=@Region" + Environment.NewLine); objParameter.Add("@Region", DbType.String, Item.Region); }
            if (Item.DQSQtyPCT != null) { SbSql.Append(",DQSQtyPCT=@DQSQtyPCT" + Environment.NewLine); objParameter.Add("@DQSQtyPCT", DbType.String, Item.DQSQtyPCT); }
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        /*刪除系統參數檔(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除系統參數檔
        /// </summary>
        /// <param name="Item">系統參數檔成員</param>
        /// <returns>回傳異動筆數</returns>
        /// <info>Author: Admin; Date: 2021/08/12  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/12  1.00    Admin        Create
        /// </history>
        public int Delete(DatabaseObject.ProductionDB.System Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [System]" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion
    }
}
