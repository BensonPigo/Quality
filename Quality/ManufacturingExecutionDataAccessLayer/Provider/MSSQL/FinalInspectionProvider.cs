using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using Newtonsoft.Json;
using DatabaseObject;
using DatabaseObject.ViewModel.FinalInspection;
using System.Transactions;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class FinalInspectionProvider : SQLDAL, IFinalInspectionProvider
    {
        #region 底層連線
        public FinalInspectionProvider(string conString) : base(conString) { }
        public FinalInspectionProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public FinalInspection GetFinalInspection(string FinalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@ID", DbType.String, FinalInspectionID }
            };

            string sqlGetData = @"
select  ID                             ,
        POID                           ,
        InspectionStage                ,
        InspectionTimes                ,
        FactoryID                      ,
        MDivisionID                    ,
        AuditDate                      ,
        SewingLineID                   ,
        AcceptableQualityLevelsUkey    ,
        SampleSize                     ,
        AcceptQty                      ,
        FabricApprovalDoc              ,
        SealingSampleDoc               ,
        MetalDetectionDoc              ,
        GarmentWashingDoc              ,
        CheckCloseShade                ,
        CheckHandfeel                  ,
        CheckAppearance                ,
        CheckPrintEmbDecorations       ,
        CheckFiberContent              ,
        CheckCareInstructions          ,
        CheckDecorativeLabel           ,
        CheckAdicomLabel               ,
        CheckCountryofOrigion          ,
        CheckSizeKey                   ,
        Check8FlagLabel                ,
        CheckAdditionalLabel           ,
        CheckShippingMark              ,
        CheckPolytagMarketing          ,
        CheckColorSizeQty              ,
        CheckHangtag                   ,
        PassQty                        ,
        RejectQty                      ,
        BAQty                          ,
        CFA                            ,
        ProductionStatus               ,
        InspectionResult               ,
        ShipmentStatus                 ,
        OthersRemark                   ,
        SubmitDate                     ,
        InspectionStep                 ,
        AddName                        ,
        AddDate                        ,
        EditName                       ,
        EditDate
from FinalInspection with (nolock)
where   ID = @ID
";
            IList<FinalInspection> listResult = ExecuteList<FinalInspection>(CommandType.Text, sqlGetData, objParameter);

            if (listResult.Count > 0)
            {
                listResult[0].Result = true;
                return listResult[0];
            }
            else
            {
                return new FinalInspection() { 
                    Result = false,
                    ErrorMessage = "No Data Found"
                };
            }
            
        }

        public string GetInspectionTimes(string POID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@POID", DbType.String, POID }
            };

            string sqlGetData = @"
select [InspectionTimes] = isnull(max(InspectionTimes), 0) + 1
    from FinalInspection
    where   POID = @POID
";
            
            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter).Rows[0]["InspectionTimes"].ToString();
        }

        private string GetNewFinalInspectionID(string factoryID)
        {
            string idHead = $"{factoryID}CH{DateTime.Now.ToString("yyMM")}";

            string sqlGetCurMaxID = $@"
select  [MaxSerID] =  cast(Replace(isnull(MAX(ID), '0'), '{idHead}', '') as int)
from    FinalInspection with (nolock)
where   ID like '{idHead}%'
";
            int newSer = (int)ExecuteDataTable(CommandType.Text, sqlGetCurMaxID, new SQLParameterCollection()).Rows[0]["MaxSerID"] + 1;
            string newID = idHead + newSer.ToString().PadLeft(4, '0');
            return newID;
        }

        public string UpdateFinalInspection(Setting setting, string userID, string factoryID, string MDivisionid)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string sqlUpdCmd = string.Empty;

            if (string.IsNullOrEmpty(setting.FinalInspectionID))
            {
                setting.FinalInspectionID = GetNewFinalInspectionID(factoryID);
                sqlUpdCmd += $@"
insert into FinalInspection(id                            ,
                            poid                          ,
                            InspectionStage               ,
                            InspectionTimes               ,
                            FactoryID                     ,
                            MDivisionid                   ,
                            AuditDate                     ,
                            SewingLineID                  ,
                            AcceptableQualityLevelsUkey   ,
                            SampleSize                    ,
                            AcceptQty                     ,
                            InspectionResult              ,
                            ShipmentStatus                ,
                            InspectionStep                ,
                            AddName                       ,
                            AddDate)
                values(@FinalInspectionID                            ,
                       @POID                          ,
                       @InspectionStage               ,
                       @InspectionTimes               ,
                       @FactoryID                     ,
                       @MDivisionid                   ,
                       @AuditDate                     ,
                       @SewingLineID                  ,
                       @AcceptableQualityLevelsUkey   ,
                       @SampleSize                    ,
                       @AcceptQty                     ,
                       'On-going'              ,
                        'On Hold'                ,
                       'Setting',
                       @UserID                       ,
                       GetDate()
                )

";
            }
            else
            {
                sqlUpdCmd += $@"
update  FinalInspection
set     InspectionStage = @InspectionStage                         ,
        AuditDate = @AuditDate    ,
        SewingLineID = @SewingLineID                             ,
        AcceptableQualityLevelsUkey = @AcceptableQualityLevelsUkey              ,
        SampleSize = @SampleSize      ,
        AcceptQty = @AcceptQty          ,
        InspectionStep = 'Setting',
        EditName = @UserID                  ,
        EditDate= getdate()
where   ID = @FinalInspectionID

delete  FinalInspection_Order where ID = @FinalInspectionID
delete  FinalInspection_OrderCarton where ID = @FinalInspectionID
";
            }



            objParameter.Add("@FinalInspectionID", setting.FinalInspectionID);
            objParameter.Add("@POID", setting.SelectedPO[0].POID);
            objParameter.Add("@InspectionStage", setting.InspectionStage);
            objParameter.Add("@InspectionTimes", setting.InspectionTimes);
            objParameter.Add("@FactoryID", factoryID);
            objParameter.Add("@MDivisionid", MDivisionid);
            objParameter.Add("@AuditDate", setting.AuditDate);
            objParameter.Add("@SewingLineID", setting.SewingLineID);
            objParameter.Add("@AcceptableQualityLevelsUkey", setting.AcceptableQualityLevelsUkey);
            objParameter.Add("@SampleSize", setting.SampleSize);
            objParameter.Add("@AcceptQty", setting.AcceptQty);
            objParameter.Add("@UserID", userID);

            foreach (SelectedPO selectedPOItem in setting.SelectedPO)
            {
                sqlUpdCmd += $@"
insert into FinalInspection_Order(ID, OrderID, AvailableQty)
            values(@FinalInspectionID, '{selectedPOItem.OrderID}', '{selectedPOItem.AvailableQty}')
";
            }

            foreach (SelectCarton selectCartonItem in setting.SelectCarton)
            {
                sqlUpdCmd += $@"
insert into FinalInspection_OrderCarton(ID, OrderID, PackingListID, CTNNo)
            values(@FinalInspectionID, '{selectCartonItem.OrderID}', '{selectCartonItem.PackingListID}', '{selectCartonItem.CTNNo}')
";
            }
            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
                transaction.Complete();
            }

            return setting.FinalInspectionID;
        }

        #endregion
    }
}
