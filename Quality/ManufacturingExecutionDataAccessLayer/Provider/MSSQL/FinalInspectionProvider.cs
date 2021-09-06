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
using System.Linq;
using DatabaseObject.RequestModel;
using System.Threading;

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
                return new FinalInspection()
                {
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
from    ManufacturingExecution.dbo.FinalInspection with (nolock)
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
delete  FinalInspection_Order_QtyShip where ID = @FinalInspectionID
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

            foreach (SelectOrderShipSeq selectOrderShipSeq in setting.SelectOrderShipSeq)
            {
                sqlUpdCmd += $@"
insert into FinalInspection_Order_QtyShip(ID, OrderID, Seq, ShipmodeID)
            values(@FinalInspectionID, '{selectOrderShipSeq.OrderID}', '{selectOrderShipSeq.Seq}', '{selectOrderShipSeq.ShipmodeID}')
";
            }

            foreach (SelectCarton selectCartonItem in setting.SelectCarton)
            {
                sqlUpdCmd += $@"
insert into FinalInspection_OrderCarton(ID, OrderID, PackingListID, CTNNo, Seq)
            values(@FinalInspectionID, '{selectCartonItem.OrderID}', '{selectCartonItem.PackingListID}', '{selectCartonItem.CTNNo}', '{selectCartonItem.Seq}')
";
            }
            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
                transaction.Complete();
            }

            return setting.FinalInspectionID;
        }

        public void UpdateFinalInspectionByStep(FinalInspection finalInspection, string currentStep, string userID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string sqlUpdCmd = string.Empty;

            switch (currentStep)
            {
                case "Setting":
                    break;
                case "Insp-General":
                    sqlUpdCmd += $@"
update FinalInspection
 set    FabricApprovalDoc = @FabricApprovalDoc  ,
        SealingSampleDoc= @SealingSampleDoc    ,
        MetalDetectionDoc= @MetalDetectionDoc   ,
        GarmentWashingDoc= @GarmentWashingDoc   ,
        InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @FinalInspectionID
";
                    objParameter.Add("@FinalInspectionID", finalInspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", finalInspection.InspectionStep);
                    objParameter.Add("@GarmentWashingDoc", finalInspection.GarmentWashingDoc);
                    objParameter.Add("@MetalDetectionDoc", finalInspection.MetalDetectionDoc);
                    objParameter.Add("@SealingSampleDoc", finalInspection.SealingSampleDoc);
                    objParameter.Add("@FabricApprovalDoc", finalInspection.FabricApprovalDoc);
                    break;
                case "Insp-CheckList":
                    sqlUpdCmd += $@"
update FinalInspection
 set    CheckCloseShade = @CheckCloseShade  ,
        CheckHandfeel = @CheckHandfeel  ,
        CheckAppearance = @CheckAppearance  ,
        CheckPrintEmbDecorations = @CheckPrintEmbDecorations  ,
        CheckFiberContent = @CheckFiberContent  ,
        CheckCareInstructions = @CheckCareInstructions  ,
        CheckDecorativeLabel = @CheckDecorativeLabel  ,
        CheckAdicomLabel = @CheckAdicomLabel  ,
        CheckCountryofOrigion = @CheckCountryofOrigion  ,
        CheckSizeKey = @CheckSizeKey  ,
        Check8FlagLabel = @Check8FlagLabel  ,
        CheckAdditionalLabel = @CheckAdditionalLabel  ,
        CheckShippingMark = @CheckShippingMark  ,
        CheckPolytagMarketing = @CheckPolytagMarketing  ,
        CheckColorSizeQty = @CheckColorSizeQty  ,
        CheckHangtag = @CheckHangtag  ,
        InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @FinalInspectionID
";

                    objParameter.Add("@FinalInspectionID", finalInspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", finalInspection.InspectionStep);
                    objParameter.Add("@CheckCloseShade", finalInspection.CheckCloseShade);
                    objParameter.Add("@CheckHandfeel", finalInspection.CheckHandfeel);
                    objParameter.Add("@CheckAppearance", finalInspection.CheckAppearance);
                    objParameter.Add("@CheckPrintEmbDecorations", finalInspection.CheckPrintEmbDecorations);
                    objParameter.Add("@CheckFiberContent", finalInspection.CheckFiberContent);
                    objParameter.Add("@CheckCareInstructions", finalInspection.CheckCareInstructions);
                    objParameter.Add("@CheckDecorativeLabel", finalInspection.CheckDecorativeLabel);
                    objParameter.Add("@CheckAdicomLabel", finalInspection.CheckAdicomLabel);
                    objParameter.Add("@CheckCountryofOrigion", finalInspection.CheckCountryofOrigion);
                    objParameter.Add("@CheckSizeKey", finalInspection.CheckSizeKey);
                    objParameter.Add("@Check8FlagLabel", finalInspection.Check8FlagLabel);
                    objParameter.Add("@CheckAdditionalLabel", finalInspection.CheckAdditionalLabel);
                    objParameter.Add("@CheckShippingMark", finalInspection.CheckShippingMark);
                    objParameter.Add("@CheckPolytagMarketing", finalInspection.CheckPolytagMarketing);
                    objParameter.Add("@CheckColorSizeQty", finalInspection.CheckColorSizeQty);
                    objParameter.Add("@CheckHangtag", finalInspection.CheckHangtag);
                    break;
                case "Insp-AddDefect":
                    break;
                case "Insp-BA":
                    break;
                case "Insp-Moisture":
                    sqlUpdCmd += $@"
update FinalInspection
 set    InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @FinalInspectionID
";
                    objParameter.Add("@FinalInspectionID", finalInspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", finalInspection.InspectionStep);
                    break;
                case "Insp-Measurement":
                    sqlUpdCmd += $@"
update FinalInspection
 set    InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @FinalInspectionID
";
                    objParameter.Add("@FinalInspectionID", finalInspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", finalInspection.InspectionStep);
                    break;
                case "Insp-Others":
                    sqlUpdCmd += $@"
update FinalInspection
 set    ProductionStatus = @ProductionStatus  ,
        OthersRemark= @OthersRemark    ,
        CFA= @CFA   ,
        InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @FinalInspectionID
";
                    objParameter.Add("@FinalInspectionID", finalInspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", finalInspection.InspectionStep);
                    objParameter.Add("@ProductionStatus", finalInspection.ProductionStatus);
                    objParameter.Add("@OthersRemark", finalInspection.OthersRemark);
                    objParameter.Add("@CFA", finalInspection.CFA);
                    break;
                case "Submit":
                    sqlUpdCmd += $@"
update FinalInspection
 set    ProductionStatus = @ProductionStatus  ,
        OthersRemark= @OthersRemark    ,
        CFA= @CFA   ,
        InspectionResult= @InspectionResult   ,
        ShipmentStatus= @ShipmentStatus   ,
        InspectionStep = 'Insp-Others',
        EditName= @userID,
        EditDate= getdate(),
        SubmitDate=getdate()
where   ID = @FinalInspectionID
";
                    objParameter.Add("@FinalInspectionID", finalInspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionResult", finalInspection.InspectionResult);
                    objParameter.Add("@ShipmentStatus", finalInspection.ShipmentStatus);
                    objParameter.Add("@ProductionStatus", finalInspection.ProductionStatus);
                    objParameter.Add("@OthersRemark", finalInspection.OthersRemark);
                    objParameter.Add("@CFA", finalInspection.CFA);
                    break;
                default:
                    break;
            }

            ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }

        public IList<byte[]> GetFinalInspectionDefectImage(long FinalInspection_DetailUkey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@FinalInspection_DetailUkey", DbType.Int64, FinalInspection_DetailUkey }
            };

            string sqlGetData = @"
select  Image
    from FinalInspection_DetailImage with (nolock)
    where   FinalInspection_DetailUkey = @FinalInspection_DetailUkey
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().Select(s => (byte[])s["Image"]).ToList();
            }
            else
            {
                return new List<byte[]>();
            }
        }

        public void UpdateFinalInspectionDetail(AddDefect addDefect, string UserID)
        {
            using (TransactionScope transaction = new TransactionScope())
            {
                SQLParameterCollection objParameter = new SQLParameterCollection() {
                    { "@FinalInspectionID", DbType.String, addDefect.FinalInspectionID },
                    { "@RejectQty", DbType.Int32, addDefect.RejectQty },
                    { "@UserID", DbType.String, UserID },
                    { "@InspectionStep", DbType.String, addDefect.InspectionStep }
                };

                string sqlUpdFinalInspection = @"
update  FinalInspection
        set PassQty = SampleSize - @RejectQty,
            RejectQty = @RejectQty,
            InspectionStep = @InspectionStep,
            EditName = @UserID,
            EditDate = getdate()
where   ID = @FinalInspectionID
";
                ExecuteNonQuery(CommandType.Text, sqlUpdFinalInspection, objParameter);

                foreach (FinalInspectionDefectItem defectItem in addDefect.ListFinalInspectionDefectItem)
                {
                    string sqlUpdateFinalInspectionDetail = string.Empty;
                    SQLParameterCollection detailParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, addDefect.FinalInspectionID },
                            { "@GarmentDefectTypeID", DbType.String, defectItem.DefectType },
                            { "@GarmentDefectCodeID", DbType.String, defectItem.DefectCode },
                            { "@Ukey", DbType.Int64, defectItem.Ukey },
                            { "@Qty", DbType.Int32, defectItem.Qty }
                        };

                    if (defectItem.Ukey == -1)
                    {
                        sqlUpdateFinalInspectionDetail = @"
    DECLARE @FinalInspection_DetailKey table (Ukey bigint)

    insert into FinalInspection_Detail(ID, GarmentDefectTypeID, GarmentDefectCodeID, Qty)
                OUTPUT INSERTED.Ukey into @FinalInspection_DetailKey
                values(@FinalInspectionID, @GarmentDefectTypeID, @GarmentDefectCodeID, @Qty)

    select  Ukey from @FinalInspection_DetailKey
";
                        DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdateFinalInspectionDetail, detailParameter);
                        defectItem.Ukey = (long)dtResult.Rows[0]["Ukey"];
                    }
                    else
                    {
                        sqlUpdateFinalInspectionDetail = @"
    if (@Qty > 0)
    begin
        update  FinalInspection_Detail
            set Qty = @Qty
            where   Ukey = @Ukey
    end
    else
    begin
        delete  FinalInspection_Detail where   Ukey = @Ukey
    end
    
";
                        ExecuteNonQuery(CommandType.Text, sqlUpdateFinalInspectionDetail, detailParameter);
                    }

                    if (defectItem.Qty > 0)
                    {
                        foreach (byte[] image in defectItem.ListFinalInspectionDefectImage)
                        {
                            string sqlInsertFinalInspection_DetailImage = @"
    insert into FinalInspection_DetailImage(ID, FinalInspection_DetailUkey, Image)
                values(@FinalInspectionID, @FinalInspection_DetailUkey, @Image)
";
                            SQLParameterCollection imgParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, addDefect.FinalInspectionID },
                            { "@FinalInspection_DetailUkey", DbType.Int64, defectItem.Ukey },
                            { "@Image", image == null ? System.Data.SqlTypes.SqlBinary.Null : image}
                        };

                            ExecuteNonQuery(CommandType.Text, sqlInsertFinalInspection_DetailImage, imgParameter);
                        }
                    }
                }

                transaction.Complete();
            }
        }

        public IList<BACriteriaItem> GetBeautifulProductAuditForInspection(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetData = $@"
select ID, Description 
into #baseBACriteria
from    [MainServer].Production.dbo.DropDownList ddl 
where Type = 'PMS_BACriteria'
order by Seq

select  [Ukey] = isnull(fn.Ukey, -1),
        [BACriteria] = bac.ID,
        [BACriteriaDesc] = bac.Description,
        [Qty] = isnull(fn.Qty, 0),		
		[RowIndex]=ROW_NUMBER() OVER(ORDER BY bac.ID) -1
    from #baseBACriteria bac with (nolock)
    left join   FinalInspection_NonBACriteria fn on    fn.ID = @finalInspectionID and
                                                            fn.BACriteria = bac.ID


";
            return ExecuteList<BACriteriaItem>(CommandType.Text, sqlGetData, listPar);
        }

        public void UpdateBeautifulProductAudit(BeautifulProductAudit beautifulProductAudit, string UserID)
        {
            using (TransactionScope transaction = new TransactionScope())
            {
                SQLParameterCollection objParameter = new SQLParameterCollection() {
                    { "@FinalInspectionID", DbType.String, beautifulProductAudit.FinalInspectionID },
                    { "@BAQty", DbType.Int32, beautifulProductAudit.BAQty },
                    { "@UserID", DbType.String, UserID },
                    { "@InspectionStep", DbType.String, beautifulProductAudit.InspectionStep }
                };

                string sqlUpdFinalInspection = @"
update  FinalInspection
        set BAQty = @BAQty,
            InspectionStep = @InspectionStep,
            EditName = @UserID,
            EditDate = getdate()
where   ID = @FinalInspectionID
";
                ExecuteNonQuery(CommandType.Text, sqlUpdFinalInspection, objParameter);

                foreach (BACriteriaItem criteriaItem in beautifulProductAudit.ListBACriteria)
                {
                    string sqlUpdateFinalInspectionDetail = string.Empty;
                    SQLParameterCollection detailParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, beautifulProductAudit.FinalInspectionID },
                            { "@BACriteria", DbType.String, criteriaItem.BACriteria },
                            { "@Ukey", DbType.Int64, criteriaItem.Ukey },
                            { "@Qty", DbType.Int32, criteriaItem.Qty }
                        };

                    if (criteriaItem.Ukey == -1)
                    {
                        sqlUpdateFinalInspectionDetail = @"
    DECLARE @FinalInspection_NonBACriteriaKey table (Ukey bigint)

    insert into FinalInspection_NonBACriteria(ID, BACriteria, Qty)
                OUTPUT INSERTED.Ukey into @FinalInspection_NonBACriteriaKey
                values(@FinalInspectionID, @BACriteria, @Qty)

    select  Ukey from @FinalInspection_NonBACriteriaKey
";
                        DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdateFinalInspectionDetail, detailParameter);
                        criteriaItem.Ukey = (long)dtResult.Rows[0]["Ukey"];
                    }
                    else
                    {
                        sqlUpdateFinalInspectionDetail = @"
    if(@Qty > 0)
    begin
        update  FinalInspection_NonBACriteria
            set Qty = @Qty
            where   Ukey = @Ukey
    end
    else
    begin
        --數量 = 0 刪除
        delete  FinalInspection_NonBACriteria where Ukey = @Ukey
    end
    
";
                        ExecuteNonQuery(CommandType.Text, sqlUpdateFinalInspectionDetail, detailParameter);
                    }

                    //數量大於0才需要上傳圖片
                    if (criteriaItem.Qty > 0)
                    {
                        foreach (byte[] image in criteriaItem.ListBACriteriaImage)
                        {
                            string sqlInsertFinalInspection_NonBACriteriaImage = @"
    insert into FinalInspection_NonBACriteriaImage(ID, FinalInspection_NonBACriteriaUkey, Image)
                values(@FinalInspectionID, @FinalInspection_NonBACriteriaUkey, @Image)
";
                            SQLParameterCollection imgParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, beautifulProductAudit.FinalInspectionID },
                            { "@FinalInspection_NonBACriteriaUkey", DbType.Int64, criteriaItem.Ukey },
                            { "@Image", image}
                        };

                            ExecuteNonQuery(CommandType.Text, sqlInsertFinalInspection_NonBACriteriaImage, imgParameter);
                        }
                    }
                }

                transaction.Complete();
            }
        }

        public List<byte[]> GetBACriteriaImage(long FinalInspection_NonBACriteriaUkey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@FinalInspection_NonBACriteriaUkey", DbType.Int64, FinalInspection_NonBACriteriaUkey }
            };

            string sqlGetData = @"
select  Image
    from FinalInspection_NonBACriteriaImage with (nolock)
    where   FinalInspection_NonBACriteriaUkey = @FinalInspection_NonBACriteriaUkey
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().Select(s => (byte[])s["Image"]).ToList();
            }
            else
            {
                return new List<byte[]>();
            }
        }

        public IList<CartonItem> GetMoistureListCartonItem(string finalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@finalInspectionID", DbType.String, finalInspectionID }
            };

            string sqlGetMoistureListCartonItem = @"
select  [FinalInspection_OrderCartonUkey] = Ukey,
        OrderID,
        PackinglistID,
        CTNNo
from    FinalInspection_OrderCarton with (nolock)
where ID = @finalInspectionID

";
            return ExecuteList<CartonItem>(CommandType.Text, sqlGetMoistureListCartonItem, objParameter);

        }

        public IList<ViewMoistureResult> GetViewMoistureResult(string finalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@finalInspectionID", DbType.String, finalInspectionID }
            };

            string sqlGetViewMoistureResult = @"
declare @FinalInspection_CTNMoisureStandard numeric(5,2)

select  @FinalInspection_CTNMoisureStandard = FinalInspection_CTNMoisureStandard
from    [MainServer].Production.dbo.System

select  fm.Ukey,
        fm.Article,
        fo.CTNNo,
        fm.Instrument,
        fm.Fabrication,
        [GarmentStandard] = em.Standard,
        fm.GarmentTop,
        fm.GarmentMiddle,
        fm.GarmentBottom,
        [CTNStandard] = @FinalInspection_CTNMoisureStandard,
        fm.CTNInside,
        fm.CTNOutside,
        fm.Result,
        fm.Action,
        fm.Remark
from    FinalInspection_Moisture fm with (nolock)
left join   FinalInspection_OrderCarton fo with (nolock) on fo.Ukey = fm.FinalInspection_OrderCartonUkey
left join   EndlineMoisture em with (nolock) on em.Instrument = fm.Instrument and em.Fabrication = fm.Fabrication
where fm.ID = @finalInspectionID

";
            return ExecuteList<ViewMoistureResult>(CommandType.Text, sqlGetViewMoistureResult, objParameter);
        }

        public IList<EndlineMoisture> GetEndlineMoisture()
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlGetEndlineMoisture = @"
select  Instrument
        ,Fabrication
        ,Standard
        ,Unit
        ,Junk
        ,Description
        ,AddDate
        ,AddName
        ,EditDate
        ,Editname
from    EndlineMoisture with (nolock)

";
            return ExecuteList<EndlineMoisture>(CommandType.Text, sqlGetEndlineMoisture, objParameter);
        }

        public void UpdateMoisture(MoistureResult moistureResult)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@FinalInspectionID", moistureResult.FinalInspectionID);
            objParameter.Add("@Article", moistureResult.Article);
            objParameter.Add("@FinalInspection_OrderCartonUkey", moistureResult.FinalInspection_OrderCartonUkey);
            objParameter.Add("@Instrument", moistureResult.Instrument);
            objParameter.Add("@Fabrication", moistureResult.Fabrication);
            objParameter.Add("@GarmentTop", moistureResult.GarmentTop);
            objParameter.Add("@GarmentMiddle", moistureResult.GarmentMiddle);
            objParameter.Add("@GarmentBottom", moistureResult.GarmentBottom);
            objParameter.Add("@CTNInside", moistureResult.CTNInside);
            objParameter.Add("@CTNOutside", moistureResult.CTNOutside);
            objParameter.Add("@Result", moistureResult.Result);
            objParameter.Add("@Action", moistureResult.Action);
            objParameter.Add("@Remark", moistureResult.Remark);
            objParameter.Add("@AddName", moistureResult.AddName);

            string sqlInsertFinalInspection_Moisture = @"
insert into FinalInspection_Moisture
(
ID
,Article
,FinalInspection_OrderCartonUkey
,Instrument
,Fabrication
,GarmentTop
,GarmentMiddle
,GarmentBottom
,CTNInside
,CTNOutside
,Result
,Action
,Remark
,AddName
,AddDate
)
values
(
 @FinalInspectionID
,@Article
,@FinalInspection_OrderCartonUkey
,@Instrument
,@Fabrication
,@GarmentTop
,@GarmentMiddle
,@GarmentBottom
,@CTNInside
,@CTNOutside
,iif(@Result = 'Pass', '1', '0')
,@Action
,@Remark
,@AddName
,GetDate()
)
";
            ExecuteNonQuery(CommandType.Text, sqlInsertFinalInspection_Moisture, objParameter);
        }

        public void DeleteMoisture(long ukey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@Ukey", ukey);

            string sqlDeleteMoisture = @" delete FinalInspection_Moisture where Ukey = @Ukey";

            ExecuteNonQuery(CommandType.Text, sqlDeleteMoisture, objParameter);
        }

        public bool CheckMoistureExists(string finalInspectionID, string article, long? finalInspection_OrderCartonUkey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string where = string.Empty;
            objParameter.Add("@FinalInspectionID", finalInspectionID);

            if (!string.IsNullOrEmpty(article))
            {
                where += " and Article = @article";
                objParameter.Add("@article", article);
            }

            if (finalInspection_OrderCartonUkey != null)
            {
                where += " and FinalInspection_OrderCartonUkey = @finalInspection_OrderCartonUkey";
                objParameter.Add("@finalInspection_OrderCartonUkey", finalInspection_OrderCartonUkey);
            }
                

            string sqlCheckMoistureExists = $@"
select  [result] = 1 
from    FinalInspection_Moisture with (nolock)
where   ID = @FinalInspectionID {where}
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlCheckMoistureExists, objParameter);

            return dtResult.Rows.Count > 0;
        }

        public void UpdateMeasurement(DatabaseObject.ViewModel.FinalInspection.Measurement measurement, string userID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            DataTable dtDateTime = ExecuteDataTableByServiceConn(CommandType.Text, "select [DateTime] = getdate()", objParameter);

            using (TransactionScope transactionScope = new TransactionScope())
            {
                foreach (MeasurementItem measurementItem in measurement.ListMeasurementItem)
                {
                    objParameter = new SQLParameterCollection();
                    objParameter.Add("@ID", measurement.FinalInspectionID);
                    objParameter.Add("@Article", measurement.SelectedArticle);
                    objParameter.Add("@SizeCode", measurement.SelectedSize);
                    objParameter.Add("@Location", measurement.SelectedProductType);
                    objParameter.Add("@Code", measurementItem.Code);
                    objParameter.Add("@SizeSpec", measurementItem.ResultSizeSpec);
                    objParameter.Add("@MeasurementUkey", measurementItem.MeasurementUkey);
                    objParameter.Add("@AddName", userID);
                    objParameter.Add("@AddDate", dtDateTime.Rows[0]["DateTime"]);

                    string sqlInsertMeasurement = @"
insert into FinalInspection_Measurement(
ID
,Article
,SizeCode
,Location
,Code
,SizeSpec
,MeasurementUkey
,AddName
,AddDate
)
values
(
 @ID
,@Article
,@SizeCode
,@Location
,@Code
,@SizeSpec
,@MeasurementUkey
,@AddName
,@AddDate
)
";
                    ExecuteNonQuery(CommandType.Text, sqlInsertMeasurement, objParameter);
                }

                transactionScope.Complete();
            }
        }

        public IList<MeasurementViewItem> GetMeasurementViewItem(string finalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@finalInspectionID", finalInspectionID);

            string sqlGetEndlineMoisture = @"
select  distinct    Article,
                    [Size] = SizeCode,
                    [ProductType] = Location,
                    [MeasurementDataByJson] = ''
from    FinalInspection_Measurement with (nolock)
where   ID = @finalInspectionID

";
            return ExecuteList<MeasurementViewItem>(CommandType.Text, sqlGetEndlineMoisture, objParameter);
        }

        public DataTable GetMeasurement(string finalInspectionID, string article, string size, string productType)
        {

            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@finalInspectionID", DbType.String, finalInspectionID} ,
                { "@article", DbType.String, article} ,
                { "@size", DbType.String, size} ,
                { "@productType", DbType.String, productType} ,
            };


            string sqlcmd = @"

declare @styleukey bigint
declare @sizeUnit varchar(8)
declare @POID varchar(13)

select  @POID = POID
from    FinalInspection with (nolock)
where   ID = @finalInspectionID

select  @styleukey = Ukey,
        @sizeUnit = SizeUnit
from    SciProduction_Style
where   Ukey = (select StyleUkey from SciProduction_Orders where ID = @POID)


select  SizeSpec,
        MeasurementUkey,
        AddDate
into    #tmp_Inspection_Measurement
from    FinalInspection_Measurement
where   ID = @finalInspectionID and
        Article = @article and
        SizeCode = @size and
        Location = @productType


select m.Ukey
	,Description = iif(isnull(b.DescEN,'') = '', m.Description, b.DescEN)
	,m.Tol1
	,m.Tol2
	,m.Code
	,m.SizeCode 
	,[MeasurementSizeSpec] = m.SizeSpec 
	,[InspectionMeasurementSizeSpec] = im.SizeSpec
	,[diff]= max(dbo.calculateSizeSpec(m.SizeSpec,im.SizeSpec, @sizeUnit))
	,im.AddDate
	,[HeadSizeCode] = FORMAT(im.AddDate,'yyyy/MM/dd HH:mm:ss')
into #tmp 
from Measurement m with(nolock)
left join #tmp_Inspection_Measurement im on im.MeasurementUkey = m.Ukey 
LEFT JOIN [ManufacturingExecution].[dbo].[MeasurementTranslate] b ON  m.MeasurementTranslateUkey = b.UKey
where   m.StyleUkey = @styleukey and
        m.SizeCode = @size and
        m.junk = 0
group by m.Ukey,iif(isnull(b.DescEN,'') = '',m.Description,b.DescEN),m.Tol1,m.Tol2,m.Code,m.SizeCode,m.SizeSpec,im.SizeSpec,im.AddDate

drop table #tmp_Inspection_Measurement

declare @HeadSizeCode as varchar(20),@mSizeCode as varchar(10),@r_id as varchar(10)
declare @sql varchar(max) = ''
DECLARE CURSOR_ CURSOR FOR
Select t.HeadSizeCode, t.SizeCode, ROW_NUMBER() over( order by t.HeadSizeCode) r_id
from #tmp t
where t.HeadSizeCode is not null
group by t.HeadSizeCode, t.SizeCode

OPEN CURSOR_
FETCH NEXT FROM CURSOR_ INTO @HeadSizeCode,@mSizeCode,@r_id
While @@FETCH_STATUS = 0
Begin
	
	set @sql = @sql + '
		,Max(case when SizeCode ='''+@mSizeCode+''' then MeasurementSizeSpec end) as ['+@mSizeCode+'_aa]
		,Max(case when HeadSizeCode ='''+@HeadSizeCode+''' and SizeCode ='''+@mSizeCode+''' then InspectionMeasurementSizeSpec end) as ['+@HeadSizeCode+']
		,Max(case when HeadSizeCode ='''+@HeadSizeCode+''' and SizeCode ='''+@mSizeCode+''' then diff end) as diff'+@r_id+''
FETCH NEXT FROM CURSOR_ INTO @HeadSizeCode,@mSizeCode,@r_id
End
CLOSE CURSOR_
DEALLOCATE CURSOR_ 

set @sql = '
	select t.Code,t.Description
		,[Tol(+)] = t.Tol2 
		,[Tol(-)] = t.Tol1 
		' + @sql + '
	from #tmp t 
	group by t.Description,t.Tol1,t.Tol2,t.code
    order by t.Code
'

exec (@sql)


drop table #tmp

";

            DataTable dt = ExecuteDataTable(CommandType.Text, sqlcmd, objParameter);
            return dt;
        }

        public List<byte[]> GetOthersImage(string finalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@finalInspectionID", DbType.String, finalInspectionID }
            };

            string sqlGetData = @"
select  Image
    from FinalInspection_OtherImage with (nolock)
    where   ID = @finalInspectionID
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().Select(s => (byte[])s["Image"]).ToList();
            }
            else
            {
                return new List<byte[]>();
            }
        }

        public void UpdateFinalInspection_OtherImage(string finalInspectionID, List<byte[]> images)
        {
            foreach (byte[] image in images)
            {
                string sqlFinalInspection_OtherImage = @"
    insert into FinalInspection_OtherImage(ID, Image)
                values(@FinalInspectionID, @Image)
";
                SQLParameterCollection imgParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, finalInspectionID },
                            { "@Image", image}
                        };

                ExecuteNonQuery(CommandType.Text, sqlFinalInspection_OtherImage, imgParameter);
            }
        }

        public DataTable GetReportMailInfo(string finalInspectionID)
        {
            SQLParameterCollection parameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, finalInspectionID }
                        };

            string sqlGetData = @"
declare @StyleID varchar(15)
declare @SeasonID varchar(10)
declare @BrandID varchar(8)

select  @StyleID = StyleID,
        @SeasonID = SeasonID,
        @BrandID = BrandID
from    SciProduction_Orders with (nolock)
where   ID = (select POID from FinalInspection with (nolock) where ID = @FinalInspectionID )

select  f.POID,
        f.InspectionResult,
        f.FactoryID,
        [SP] = (SELECT Stuff((select concat( ',',OrderID) 
                                from  FinalInspection_Order fo with (nolock) 
                                where fo.ID = f.ID
                                FOR XML PATH('')),1,1,'') ),
        [StyleID] = @StyleID,
        [SeasonID] = @SeasonID,
        [BrandID] = @BrandID,
        [CFA] = (select name from pass1 with (nolock) where ID = f.CFA),
        [SubmitDate] = Format(f.SubmitDate, 'yyyy/MM/dd'),
        [AuditDate] = Format(f.AuditDate, 'yyyy/MM/dd')
from FinalInspection f with (nolock)
where   f.ID = @FinalInspectionID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, parameter);
        }

        public DataTable GetQueryReportInfo(string finalInspectionID)
        {
            SQLParameterCollection parameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, finalInspectionID }
                        };

            string sqlGetData = @"
declare @StyleID varchar(15)
declare @SeasonID varchar(10)
declare @BrandID varchar(8)
declare @spQty int

select @spQty = isnull(sum(Qty), 0)
from SciProduction_Orders with (nolock)
where   ID in (select OrderID from FinalInspection_Order with (nolock) where ID = @FinalInspectionID)

select  @StyleID = StyleID,
        @SeasonID = SeasonID,
        @BrandID = BrandID
from    SciProduction_Orders with (nolock)
where   ID = (select POID from FinalInspection with (nolock) where ID = @FinalInspectionID )

select  [SP] = (SELECT Stuff((select concat( ',',OrderID) 
                                from  FinalInspection_Order fo with (nolock) 
                                where fo.ID = @FinalInspectionID
                                FOR XML PATH('')),1,1,'') ),
        [StyleID] = @StyleID,
        [BrandID] = @BrandID,
        [CFA] = (select name from pass1 with (nolock) where ID = f.CFA),
        [TotalSPQty] = @spQty
from    FinalInspection f with (nolock)
where ID = @FinalInspectionID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, parameter);
        }

        public IList<QueryFinalInspection> GetFinalinspectionQueryList(FinalInspection_Query request)
        {
            SQLParameterCollection parameter = new SQLParameterCollection();

            string whereOrder = string.Empty;
            string whereFinalInspection = string.Empty;

            if (!string.IsNullOrEmpty(request.InspectionResult))
            {
                whereFinalInspection += " and f.InspectionResult = @InspectionResult ";
                whereOrder += @" and ID in (select  distinct OrderID
from FinalInspection f with (nolock)
inner join  FinalInspection_Order fo with (nolock) on fo.ID = f.ID
where f.InspectionResult = @InspectionResult)";
                parameter.Add("@InspectionResult", request.InspectionResult);
            }

            if (!string.IsNullOrEmpty(request.POID))
            {
                whereOrder += @" and POID = @POID";
                whereFinalInspection += @" and f.POID = @POID";
                parameter.Add("@POID", request.POID);
            }

            if (!string.IsNullOrEmpty(request.SP))
            {
                whereOrder += @" and ID = @SP";
                whereFinalInspection += @" and fo.OrderID = @SP";
                parameter.Add("@SP", request.SP);
            }

            if (request.SciDeliveryStart != null)
            {
                whereOrder += @" and SciDelivery >= @SciDeliveryStart";
                parameter.Add("@SciDeliveryStart", request.SciDeliveryStart);
            }

            if (request.SciDeliveryEnd != null)
            {
                whereOrder += @" and SciDelivery <= @SciDeliveryEnd";
                parameter.Add("@SciDeliveryEnd", request.SciDeliveryEnd);
            }

            if (!string.IsNullOrEmpty(request.StyleID))
            {
                whereOrder += @" and StyleID = @StyleID";
                parameter.Add("@StyleID", request.StyleID);
            }

            string sqlGetData = $@"

select  ID,
        StyleID,
        SeasonID,
        BrandID,
        Qty
into    #tmpOrders
from    SciProduction_Orders with (nolock)
where   1 = 1 {whereOrder}

select  ID, Article
into    #tmpOrderArticle
from    [MainServer].Production.dbo.Order_Article with (nolock)
where   ID in (select ID from #tmpOrders)

select  [FinalInspectionID] = f.ID,
        [SP] = fo.OrderID,
        f.POID,
        [SPQty] = cast(o.Qty as varchar),
        [StyleID] = o.StyleID,
        [Season] = o.SeasonID,
        [BrandID] = o.BrandID,
        [Article] = (SELECT Stuff((select concat( ',',Article)   from #tmpOrderArticle where ID = fo.OrderID FOR XML PATH('')),1,1,'') ),
        [InspectionTimes] = cast(f.InspectionTimes as varchar),
        f.InspectionStage,
        f.InspectionResult
from FinalInspection f with (nolock)
inner join  FinalInspection_Order fo with (nolock) on fo.ID = f.ID
inner join  #tmpOrders o on fo.OrderID = o.ID
where   1 = 1 {whereFinalInspection}
";

            return ExecuteList<QueryFinalInspection>(CommandType.Text, sqlGetData, parameter);
        }

        public IList<FinalInspection_OrderCarton> GetListCartonInfo(string finalInspectionID)
        {
            SQLParameterCollection parameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, finalInspectionID }
                        };

            string sqlGetData = @"
select  Ukey
        ,ID
        ,OrderID
        ,PackinglistID
        ,CTNNo
from    FinalInspection_OrderCarton with (nolock)
where   ID = @FinalInspectionID
";

            return ExecuteList<FinalInspection_OrderCarton>(CommandType.Text, sqlGetData, parameter);
        }

        #endregion
    }
}
