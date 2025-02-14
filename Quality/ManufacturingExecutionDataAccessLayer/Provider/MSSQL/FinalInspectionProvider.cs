using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.Public;
using DatabaseObject.ViewModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Transactions;
using ToolKit;

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
        BrandID = brand.BrandID ,--(select top 1 BrandID from SciProduction_Orders o with (nolock) where o.CustPONO = a.CustPONO)                       ,
        [CustPONO] = o.val             ,
        InspectionStage                ,
        InspectionTimes                ,
        FactoryID                      ,
        MDivisionID                    ,
        AuditDate                      ,
        SewingLineID                   

        ,AcceptableQualityLevelsUkey    
        ,AcceptableQualityLevelsProUkey    

        ,SampleSize                     
        ,AcceptQty                      
        ,FabricApprovalDoc              
        ,SealingSampleDoc               
        ,MetalDetectionDoc              ,
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
        CheckPolytagMarking          ,
        CheckColorSizeQty              ,
        CheckHangtag                   ,
        PassQty                        ,
        RejectQty      ,
        BAQty                          ,
        CFA                            ,
        Clerk                            ,
        ProductionStatus =  a.ProductionStatus              ,
        ProductionStatusDefault = Cast( ISNULL(ProductionStatusVal.val,0) as decimal)             ,
        InspectionResult               ,
        ShipmentStatus                 ,
        OthersRemark                   ,
        SubmitDate                     ,
        InspectionStep                 ,
        Shift                     ,
        Team                 ,
        AddName                        ,
        AddDate                        ,
        EditName                       ,
        EditDate                       ,
        ReInspection                       ,
        P88UniqueKey                       ,
        IsExportToP88  ,
        IsFollowAQL,
        HasOtherImage = Cast(IIF(exists(select 1 from SciPMSFile_FinalInspection_OtherImage b WITH(NOLOCK) where a.id= b.id),1,0) as bit),
        CheckFGPT                      ,
        [FGWT] = iif(a.InspectionStage in ('Final' ,'Final Internal'), ISNULL(g.WashResult, 'Lacking Test') , ''),
        [FGPT] = iif(a.InspectionStage in ('Final' ,'Final Internal'), fgpt.Result, ''),
        [ISFD] = cast(I.ISFD as bit) 
from FinalInspection a with (nolock)
outer apply (
	select TOP 1 o.BrandID
	from FinalInspection_Order fo
	left join Production.dbo.Orders o with (nolock) on fo.OrderID = o.ID
	where fo.ID = a.ID
)brand
outer apply (
    SELECT val = Stuff((select distinct concat( ',',CustPONo) 
                       from  FinalInspection_Order fo with (nolock) 
                       inner join SciProduction_Orders o with (nolock) on fo.OrderID = o.ID
                       where fo.ID = a.ID
                       FOR XML PATH('')),1,1,'')
)o
outer apply (
	select [GarmentTestID] = g.ID, [WashResult] = case g.WashResult when 'F' then 'Failed Test' when 'P' then 'Completed Test' else 'Lacking Test' end
	from FinalInspection_Order o 
	left join SciProduction_GarmentTest g on o.OrderID = g.OrderID
	where o.ID = a.ID
)g
outer apply (
	select [Result] = case MAX(case [Result] when 'fail' then 2 when 'pass' then 1 else 0 end)  when 2 then 'Failed Test' when 1 then 'Completed Test' else 'Lacking Test' end
	from (
		select distinct [Result] = Lower(CASE WHEN  t.TestUnit = 'N' AND t.[TestResult] !='' THEN IIF( Cast( t.[TestResult] as float) >= cast( t.Criteria as float) ,'Pass' ,'Fail')
					WHEN  t.TestUnit = 'mm' THEN IIF(  t.[TestResult] = '<=4' OR t.[TestResult] = '≦4','Pass' , IIF( t.[TestResult]='>4','Fail','')  )
					WHEN  t.TestUnit = 'Pass/Fail' THEN t.[TestResult]
			   ELSE ''
			END)
		from SciProduction_GarmentTest_Detail_FGPT t with(nolock)
		where g.[GarmentTestID] = t.ID
	)t
)fgpt
outer apply (
	select [ISFD] = 
			MAX(case when sr.RR = 1 or sr.LR = 1 then 1
				when s.ExpectionFormDate >= DATEADD(Year,-2, GETDATE()) then 1
			else 0
			end)
	from SciProduction_Orders o with (nolock)
	inner join SciProduction_Style s with (nolock) on s.Ukey = o.StyleUkey
	left join SciProduction_Style_RRLR_Report sr with (nolock) on s.Ukey = sr.StyleUkey
	where exists (select 1 from FinalInspection_Order fo with (nolock) where fo.ID = a.ID and fo.OrderID = o.ID)
)I
Outer Apply(----比照PMS QA P32：根據SP# + PackingList_Detail.OrderShipmodeSeq(運輸方式)，去撈取所有紙箱，計算Clog 收到紙箱的百分比
    SELECT Val= CAST(ROUND( SUM(IIF( CFAReceiveDate IS NOT NULL OR ReceiveDate IS NOT NULL
								    ,ShipQty
								    ,0)
						    ) * 1.0 
						    /  SUM(ShipQty) * 100 
    ,0) AS INT) 
    FROM MainServer.Production.dbo.PackingList_Detail pd WITH(NOLOCK)
    INNER JOIN ManufacturingExecution..FinalInspection_Order_QtyShip foq on pd.OrderID=foq.OrderID and pd.OrderShipmodeSeq=foq.Seq
    WHERE foq.ID =  a.ID
)ProductionStatusVal
where a.ID = @ID
";
            IList<FinalInspection> listResult = ExecuteList<FinalInspection>(CommandType.Text, sqlGetData, objParameter);

            sqlGetData = "select * from FinalInspectionGeneral where FinalInspectionID = @ID ";
            var Generaltmp = ExecuteList<FinalInspectionGeneral>(CommandType.Text, sqlGetData, objParameter);

            sqlGetData = "select * from FinalInspectionCheckList where FinalInspectionID = @ID ";
            var CheckListRTmp = ExecuteList<FinalInspectionCheckList>(CommandType.Text, sqlGetData, objParameter);


            if (listResult.Count > 0)
            {
                FinalInspection result = listResult[0];
                result.Result = true;

                result.finalInspectionGeneral = Generaltmp.Any() ? Generaltmp.FirstOrDefault() : new FinalInspectionGeneral();
                result.finalInspectionCheckList = CheckListRTmp.Any() ? CheckListRTmp.FirstOrDefault() : new FinalInspectionCheckList();

                result.finalInspectionGeneral.FinalInspectionID = result.ID;
                result.finalInspectionCheckList.FinalInspectionID = result.ID;

                return result;
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

        /// <summary>
        /// 取得可用的關卡
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="CustPONO"></param>
        /// <returns></returns>
        public List<FinalInspection_Step> GetAllStep(string FinalInspectionID, string CustPONO)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
                { "@FinalInspectionID", DbType.String, FinalInspectionID },
                { "@CustPONO", DbType.String, CustPONO }
            };

            string cmd = string.Empty;

            if (!string.IsNullOrEmpty(FinalInspectionID))
            {
                cmd = $@"
select DISTINCT FinalInspectionID = a.ID,b.BrandID,c.Seq,d.StepName,c.StepUkey
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID=b.BrandID
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
where a.ID = @FinalInspectionID
UNION----無客製，則抓預設關卡
select DISTINCT FinalInspectionID = a.ID,b.BrandID,c.Seq,d.StepName,c.StepUkey
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID='DEFAULT'
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
where a.ID = @FinalInspectionID
AND NOT EXISTS(
	select 1 from  FinalInspectionBasicBrand_Step where BrandID = b.BrandID
)
";
            }
            else
            {
                cmd = $@"
select DISTINCT FinalInspectionID = '',b.BrandID,c.Seq,d.StepName,c.StepUkey
from Production..Orders b 
inner join FinalInspectionBasicBrand_Step c on c.BrandID=b.BrandID
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
where b.CustPONO = @CustPONO
UNION----無客製，則抓預設關卡
select DISTINCT FinalInspectionID = '',b.BrandID,c.Seq,d.StepName,c.StepUkey
from Production..Orders b 
inner join FinalInspectionBasicBrand_Step c on c.BrandID='DEFAULT'
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
where b.CustPONO = @CustPONO
AND NOT EXISTS(
	select 1 from  FinalInspectionBasicBrand_Step where BrandID = b.BrandID
)

                ";
            }

            var r = ExecuteList<FinalInspection_Step>(CommandType.Text, cmd, objParameter);

            return r.Any() ? r.ToList() : new List<FinalInspection_Step>();
        }

        /// <summary>
        /// 取得Final單子的上一關or 下一關 or 目前關卡
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        public List<FinalInspection_Step> GetAllStepByAction(string FinalInspectionID, FinalInspectionSStepAction action)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
                { "@FinalInspectionID", DbType.String, FinalInspectionID },

            };

            string cmd = $@"
----該Final單的所有可用關卡
select b.BrandID,Seq = 0 ,StepName = 'Insp-Setting',StepUkey=0 ----Setting是全世界必經之路，額外處理
INTO #AllStep
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
where a.ID = @FinalInspectionID
UNION  ----品牌客製關卡
select b.BrandID,c.Seq,d.StepName,c.StepUkey
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID=b.BrandID
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
where a.ID = @FinalInspectionID and a.SubmitDate is null
UNION  ----無客製，則抓預設關卡
select b.BrandID,c.Seq,d.StepName,c.StepUkey
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID='DEFAULT'
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
where a.ID = @FinalInspectionID and a.SubmitDate is null
AND NOT EXISTS(
	select 1 from  FinalInspectionBasicBrand_Step where BrandID = b.BrandID
)

----找出現在關卡
select b.BrandID,Seq=0,StepName='Insp-Setting',StepUkey =0
INTO #CurrentStep
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
where a.ID = @FinalInspectionID AND a.InspectionStep='Insp-Setting'
UNION
select b.BrandID,c.Seq,d.StepName,c.StepUkey 
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID=b.BrandID
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
outer apply(
	select DISTINCT StepName = Data 
	from SplitString(a.InspectionStep,'-') s
	
)currentStep
where a.id = @FinalInspectionID and a.SubmitDate is null
and currentStep.StepName IN (select REPLACE(q.StepName,' ','') from FinalInspectionBasicStep q with(NOLOCK) )
and currentStep.StepName=REPLACE(d.StepName,' ','') 
UNION  ----無客製，則抓預設關卡
select b.BrandID,c.Seq,d.StepName,c.StepUkey 
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID='DEFAULT'
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
outer apply(
	select DISTINCT StepName = Data 
	from SplitString(a.InspectionStep,'-') s
	
)currentStep
where a.id = @FinalInspectionID and a.SubmitDate is null
and currentStep.StepName IN (select REPLACE(q.StepName,' ','') from FinalInspectionBasicStep q with(NOLOCK) )
and currentStep.StepName=REPLACE(d.StepName,' ','') 
AND NOT EXISTS(
	select 1 from  FinalInspectionBasicBrand_Step where BrandID = b.BrandID
)


select TOP 1 a.BrandID ,a.Seq ,StepName = REPLACE(a.StepName,' ',''),a.StepUkey
from #AllStep a
where 1=1
";

            switch (action)
            {
                case FinalInspectionSStepAction.Next:
                    cmd += $@"
----下一關卡
and Seq > (select Seq from #CurrentStep) order by Seq  
";
                    break;
                case FinalInspectionSStepAction.Previous:
                    cmd += $@"
----上一關卡
and Seq < (select Seq from #CurrentStep) order by Seq desc 
";
                    break;
                case FinalInspectionSStepAction.Current:
                    cmd += $@"
----當下關卡
and Seq = (select Seq from #CurrentStep) order by Seq  
";
                    break;
                default:
                    cmd += $@"
----當下關卡
and Seq = (select Seq from #CurrentStep) order by Seq  
";
                    break;
            }

            cmd += $@"drop table #AllStep,#CurrentStep";

            var r = ExecuteList<FinalInspection_Step>(CommandType.Text, cmd, objParameter);

            return r.Any() ? r.ToList() : new List<FinalInspection_Step>();
        }

        /// <summary>
        /// 將Final單子前往下一關/回到上一關
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="action">前往下一關/回到上一關</param>
        /// <returns></returns>
        public void UpdateStepByAction(string FinalInspectionID, string UserID, FinalInspectionSStepAction action)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
                { "@FinalInspectionID", DbType.String, FinalInspectionID },
                { "@UserID", DbType.String, UserID },

            };

            string sqlUpdCmd = $@"DECLARE @TargetStep as varchar(50) = ''

----該Final單的所有可用關卡
select b.BrandID,Seq = 0 ,StepName = 'Insp-Setting',StepUkey=0 ----Setting是全世界必經之路，額外處理
INTO #AllStep
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
where a.ID = @FinalInspectionID
UNION  ----品牌客製關卡
select b.BrandID,c.Seq,d.StepName,c.StepUkey
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID=b.BrandID
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
where a.ID = @FinalInspectionID and a.SubmitDate is null
UNION----無客製，則抓預設關卡
select b.BrandID,c.Seq,d.StepName,c.StepUkey
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID='DEFAULT'
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
where a.ID = @FinalInspectionID and a.SubmitDate is null
AND NOT EXISTS(
	select 1 from  FinalInspectionBasicBrand_Step where BrandID = b.BrandID
)


----找出現在關卡
select b.BrandID,Seq=0,StepName='Insp-Setting',StepUkey =0
INTO #CurrentStep
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
where a.ID = @FinalInspectionID AND a.InspectionStep='Insp-Setting'
UNION
select b.BrandID,c.Seq,d.StepName,c.StepUkey 
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID=b.BrandID
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
outer apply(
	select DISTINCT StepName = Data 
	from SplitString(a.InspectionStep,'-') s
	
)currentStep
where a.id = @FinalInspectionID and a.SubmitDate is null
and currentStep.StepName IN (select REPLACE(q.StepName,' ','') from FinalInspectionBasicStep q with(NOLOCK) )
and currentStep.StepName=REPLACE(d.StepName,' ','') 
UNION  ----無客製，則抓預設關卡
select b.BrandID,c.Seq,d.StepName,c.StepUkey 
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_Step c on c.BrandID='DEFAULT'
inner join FinalInspectionBasicStep d on d.Ukey = c.StepUkey
outer apply(
	select DISTINCT StepName = Data 
	from SplitString(a.InspectionStep,'-') s
	
)currentStep
where a.id = @FinalInspectionID and a.SubmitDate is null
and currentStep.StepName IN (select REPLACE(q.StepName,' ','') from FinalInspectionBasicStep q with(NOLOCK) )
and currentStep.StepName=REPLACE(d.StepName,' ','') 
AND NOT EXISTS(
	select 1 from  FinalInspectionBasicBrand_Step where BrandID = b.BrandID
)

select TOP 1 @TargetStep = 'Insp-' +  REPLACE((REPLACE(a.StepName,' ','')),'Insp-','')
from #AllStep a
where 1=1
";

            switch (action)
            {
                case FinalInspectionSStepAction.Next:
                    sqlUpdCmd += $@"
----下一關卡
and Seq > (select Seq from #CurrentStep) order by Seq  
";
                    break;
                case FinalInspectionSStepAction.Previous:
                    sqlUpdCmd += $@"
----上一關卡
and Seq < (select Seq from #CurrentStep) order by Seq desc 
";
                    break;
                case FinalInspectionSStepAction.Current:
                    sqlUpdCmd += $@"
----當下關卡
and Seq = (select Seq from #CurrentStep) order by Seq  
";
                    break;
                default:
                    sqlUpdCmd += $@"
----當下關卡
and Seq = (select Seq from #CurrentStep) order by Seq  
";
                    break;
            }

            sqlUpdCmd += $@"

IF @TargetStep=''
	SET @TargetStep= 'Submit'


UPDATE FinalInspection 
SET InspectionStep = @TargetStep
    ,EditName= @UserID
    ,EditDate= getdate()
where ID = @FinalInspectionID

drop table #AllStep,#CurrentStep
";

            ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }
        public string GetInspectionTimes(string CustPONO)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@CustPONO", DbType.String, CustPONO }
            };

            string sqlGetData = @"
select [InspectionTimes] = isnull(max(InspectionTimes), 0) + 1
    from FinalInspection WITH(NOLOCK)
    where   CustPONO = @CustPONO
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter).Rows[0]["InspectionTimes"].ToString();
        }

        public string GetNewFinalInspectionID(string factoryID)
        {
            string idHead = $"{factoryID}CH{DateTime.Now.ToString("yyMM")}";

            string sqlGetCurMaxID = $@"
select  [MaxSerID] =  cast(Replace(isnull(MAX(ID), '0'), '{idHead}', '') as int)
from    ManufacturingExecution.dbo.FinalInspection with (nolock)
where   ID like '{idHead}%'
";
            int newSer = (int)ExecuteDataTableByServiceConn(CommandType.Text, sqlGetCurMaxID, new SQLParameterCollection()).Rows[0]["MaxSerID"] + 1;
            string newID = idHead + newSer.ToString().PadLeft(4, '0');
            return newID;
        }

        public string UpdateFinalInspection(Setting setting, string userID, string factoryID, string MDivisionid, string NewFinalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string sqlUpdCmd = string.Empty;

            if (string.IsNullOrEmpty(setting.FinalInspectionID))
            {
                setting.FinalInspectionID = NewFinalInspectionID;
                sqlUpdCmd += $@"
insert into FinalInspection(id                            ,
                            CustPONO                          ,
                            InspectionStage               ,
                            InspectionTimes               ,
                            FactoryID                     ,
                            MDivisionid                   ,
                            AuditDate                     ,
                            SewingLineID                  ,
                            AcceptableQualityLevelsUkey   ,
                            AcceptableQualityLevelsProUkey   ,
                            SampleSize                    ,
                            AcceptQty                     ,
                            InspectionResult              ,
                            ShipmentStatus                ,
                            InspectionStep                ,
                            Shift                         ,
                            Team                          ,
                            AddName                       ,
                            AddDate                       ,
                            ReInspection                   ,
                            MeasurementAQLUkey,
                            MeasurementSampleSize,
                            MeasurementAcceptQty,
                            IsFollowAQL             ,
                            P88UniqueKey               
                        )
                values(@FinalInspectionID                            ,
                       @CustPONO                          ,
                       @InspectionStage               ,
                       @InspectionTimes               ,
                       @FactoryID                     ,
                       @MDivisionid                   ,
                       @AuditDate                     ,
                       @SewingLineID                  ,
                       @AcceptableQualityLevelsUkey   ,
                       @AcceptableQualityLevelsProUkey   ,
                       @SampleSize                    ,
                       @AcceptQty                     ,
                       'On-going'              ,
                        'On Hold'                ,
                       'Insp-Setting',
                       @Shift                         ,
                       @Team                          ,
                       @UserID                       ,
                       GetDate()                     ,
                       @ReInspection                   ,
                    ISNULL( (----用現用的AQL範圍，去找Measurement專用的AQL，所以要限定Category=Measurement
                        select TOP 1 t.Ukey
                        from SciProduction_AcceptableQualityLevels t
                        where t.BrandID = (select TOP 1 BrandID from SciProduction_Orders where CustPONO= @CustPONO)
	                    and t.Category='Measurement' 
	                    and EXISTS(
		                    select 1 from SciProduction_AcceptableQualityLevels s 
		                    where s.LotSize_Start=s.LotSize_Start AND t.LotSize_End=s.LotSize_End
		                    and s.Ukey = @AcceptableQualityLevelsUkey
	                    )
                    ),0) ,
                    ISNULL( (
                        select TOP 1 t.SampleSize
                        from SciProduction_AcceptableQualityLevels t
                        where t.BrandID = (select TOP 1 BrandID from SciProduction_Orders where CustPONO= @CustPONO)
	                    and t.Category='Measurement' 
	                    and EXISTS(
		                    select 1 from SciProduction_AcceptableQualityLevels s 
		                    where s.LotSize_Start=s.LotSize_Start AND t.LotSize_End=s.LotSize_End
		                    and s.Ukey = @AcceptableQualityLevelsUkey
	                    )
                    ),0) ,
                    ISNULL( (
                        select TOP 1 t.AcceptedQty
                        from SciProduction_AcceptableQualityLevels t
                        where t.BrandID = (select TOP 1 BrandID from SciProduction_Orders where CustPONO= @CustPONO)
	                    and t.Category='Measurement' 
	                    and EXISTS(
		                    select 1 from SciProduction_AcceptableQualityLevels s 
		                    where s.LotSize_Start=s.LotSize_Start AND t.LotSize_End=s.LotSize_End
		                    and s.Ukey = @AcceptableQualityLevelsUkey
	                    )
                    ),0),
                    @IsFollowAQL          
                    ,'sintex' + @FinalInspectionID  
                )
;
INSERT INTO FinalInspectionGeneral
           (FinalInspectionID)
     VALUES
           (@FinalInspectionID)
;
INSERT INTO FinalInspectionCheckList
           (FinalInspectionID)
     VALUES
           (@FinalInspectionID)
;

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
        AcceptableQualityLevelsProUkey = @AcceptableQualityLevelsProUkey              ,
        SampleSize = @SampleSize      ,
        AcceptQty = @AcceptQty          ,
        InspectionStep = 'Insp-Setting',
        Shift = @Shift          ,
        Team = @Team,
        ReInspection = @ReInspection,       
        IsFollowAQL = @IsFollowAQL,                         
        EditName = @UserID                  ,
        EditDate= getdate(),

        MeasurementAQLUkey = ISNULL( (----用現用的AQL範圍，去找Measurement專用的AQL，所以要限定Category=Measurement
                                select TOP 1 t.Ukey
                                from SciProduction_AcceptableQualityLevels t
                                where t.BrandID = (select TOP 1 BrandID from SciProduction_Orders where CustPONO= @CustPONO)
	                            and t.Category='Measurement' 
	                            and EXISTS(
		                            select 1 from SciProduction_AcceptableQualityLevels s 
		                            where s.LotSize_Start=s.LotSize_Start AND t.LotSize_End=s.LotSize_End
		                            and s.Ukey = @AcceptableQualityLevelsUkey
	                            )
                            ) ,0)             ,
        MeasurementSampleSize = ISNULL( (
                                select TOP 1 t.SampleSize
                                from SciProduction_AcceptableQualityLevels t
                                where t.BrandID = (select TOP 1 BrandID from SciProduction_Orders where CustPONO= @CustPONO)
	                            and t.Category='Measurement' 
	                            and EXISTS(
		                            select 1 from SciProduction_AcceptableQualityLevels s 
		                            where s.LotSize_Start=s.LotSize_Start AND t.LotSize_End=s.LotSize_End
		                            and s.Ukey = @AcceptableQualityLevelsUkey
	                            )
                            ) ,0)             ,
        MeasurementAcceptQty = ISNULL( (
                                select TOP 1 t.AcceptedQty
                                from SciProduction_AcceptableQualityLevels t
                                where t.BrandID = (select TOP 1 BrandID from SciProduction_Orders where CustPONO= @CustPONO)
	                            and t.Category='Measurement' 
	                            and EXISTS(
		                            select 1 from SciProduction_AcceptableQualityLevels s 
		                            where s.LotSize_Start=s.LotSize_Start AND t.LotSize_End=s.LotSize_End
		                            and s.Ukey = @AcceptableQualityLevelsUkey
	                            )
                            ) ,0)

where   ID = @FinalInspectionID

delete  FinalInspection_Order where ID = @FinalInspectionID
delete  FinalInspection_Order_QtyShip where ID = @FinalInspectionID
delete  FinalInspection_OrderCarton where ID = @FinalInspectionID
delete  FinalInspection_Order_Breakdown where FinalInspectionID = @FinalInspectionID
";
            }



            objParameter.Add("@FinalInspectionID", setting.FinalInspectionID);
            objParameter.Add("@CustPONO", setting.SelectedPO[0].CustPONO ?? string.Empty);
            objParameter.Add("@InspectionStage", setting.InspectionStage);
            objParameter.Add("@InspectionTimes", setting.InspectionTimes);
            objParameter.Add("@FactoryID", factoryID);
            objParameter.Add("@MDivisionid", MDivisionid);
            objParameter.Add("@AuditDate", setting.AuditDate);
            objParameter.Add("@SewingLineID", (setting.SewingLineID == null ? string.Empty : setting.SewingLineID));
            objParameter.Add("@AcceptableQualityLevelsUkey", setting.AcceptableQualityLevelsUkey ?? "0");
            objParameter.Add("@AcceptableQualityLevelsProUkey", setting.AcceptableQualityLevelsProUkey ?? "0");
            objParameter.Add("@SampleSize", setting.SampleSize);
            objParameter.Add("@AcceptQty", setting.AcceptQty);
            objParameter.Add("@UserID", userID);
            objParameter.Add("@Team", setting.Team);
            objParameter.Add("@Shift", setting.Shift);
            objParameter.Add("@ReInspection", setting.ReInspection);
            objParameter.Add("@IsFollowAQL", setting.IsFollowAQL);

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
insert into FinalInspection_Order_QtyShip(ID, OrderID, Seq, ShipmodeID, InspectionTimes)
            values(@FinalInspectionID, '{selectOrderShipSeq.OrderID}', '{selectOrderShipSeq.Seq}', '{selectOrderShipSeq.ShipmodeID}'
, (
	select [InspectionTimes] = ISNULL(MAX(foq.InspectionTimes), 0) + 1
	from FinalInspection_Order_QtyShip foq
	left join FinalInspection f on foq.ID = f.ID
	where exists (select 1 from FinalInspection f2 where f2.ID = @FinalInspectionID and f2.AddDate > f.AddDate)
	and foq.OrderID = '{selectOrderShipSeq.OrderID}'
	and foq.Seq = '{selectOrderShipSeq.Seq}')
)
";
            }

            foreach (SelectCarton selectCartonItem in setting.SelectCarton)
            {
                sqlUpdCmd += $@"
insert into FinalInspection_OrderCarton(ID, OrderID, PackingListID, CTNNo, Seq)
            values(@FinalInspectionID, '{selectCartonItem.OrderID}', '{selectCartonItem.PackingListID}', '{selectCartonItem.CTNNo}', '{selectCartonItem.Seq}')
";
            }

            foreach (SelectQtyBreakdown selectQtyBreakdown in setting.SelectQtyBreakdownList)
            {
                sqlUpdCmd += $@"
insert into FinalInspection_Order_Breakdown (FinalInspectionID, OrderID, Article,SizeCode, LineItem, Junk)
            values(@FinalInspectionID, '{selectQtyBreakdown.OrderID}', '{selectQtyBreakdown.Article}', '{selectQtyBreakdown.SizeCode}', '{selectQtyBreakdown.LineItem ??  0}', {(selectQtyBreakdown.Junk ? "1":"0")} ) 
";
            }

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
                transaction.Complete();
            }

            return setting.FinalInspectionID;
        }


        public void ModifyFinalInspection_Order_Breakdown(List<FinalInspection_Order_Breakdown> qtyBreakdownList, string p88UniqueKey)
        {
            string sqlUpdCmd = $@"DELETE FROM FinalInspection_Order_Breakdown WHERE FinalInspectionID = '{qtyBreakdownList.FirstOrDefault().FinalInspectionID}' AND OrderID  = '{qtyBreakdownList.FirstOrDefault().OrderID}';";
            foreach (var item in qtyBreakdownList)
            {
                sqlUpdCmd += $@"
INSERT INTO FinalInspection_Order_Breakdown (FinalInspectionID, OrderID, Article, SizeCode, LineItem, Junk)
            VALUES('{item.FinalInspectionID}', '{item.OrderID}', '{item.Article}', '{item.SizeCode}', '{item.LineItem}', {(item.Junk ? "1" : "0")} ) 
;
";
            }
            sqlUpdCmd += $@"UPDATE FinalInspection SET P88UniqueKey = dbo.GetNextP88UniqueKey(P88UniqueKey)  WHERE ID = '{qtyBreakdownList.FirstOrDefault().FinalInspectionID}'; ";
            ExecuteNonQuery(CommandType.Text, sqlUpdCmd, new SQLParameterCollection());
        }


        public void UpdateAuditDate(FinalInspection f)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@AuditDate", f.AuditDate);
            objParameter.Add("@FinalInspectionID", f.ID);
            string sqlUpdCmd = $@"
UPDATE FinalInspection
SET AuditDate = @AuditDate ,EditDate = GETDATE()
WHERE ID = @FinalInspectionID";

            ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }
        public int GetAvailableQty(string FinalInspectionID)
        {
            string sql = $@"
select AvailableQty=SUM(AvailableQty)
from FinalInspection_Order
where   ID = '{FinalInspectionID}'
";
            int AvailableQty = (int)ExecuteDataTableByServiceConn(CommandType.Text, sql, new SQLParameterCollection()).Rows[0]["AvailableQty"];
            return AvailableQty;
        }

        /// <summary>
        /// 部分功能Back/Next按鈕按下時，要存檔的東西(Remark之類的)
        /// </summary>
        /// <param name="finalInspection"></param>
        /// <param name="currentStep"></param>
        /// <param name="userID"></param>
        public void UpdateStepInfo(FinalInspection finalInspection, string currentStep, string userID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string sqlUpdCmd = string.Empty;

            // ISP20230647 開始，關卡的轉移UPDATE透過其他Function異動，這邊只保留部分功能Back/Next按鈕按下時，要存檔的東西(Remark之類的)
            switch (currentStep)
            {
                case "Insp-General":
                    sqlUpdCmd += $@"
update FinalInspection
 set    EditName= @userID,
        EditDate= getdate()
where   ID = @FinalInspectionID
";
                    objParameter.Add("@FinalInspectionID", finalInspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@GarmentWashingDoc", finalInspection.GarmentWashingDoc);
                    objParameter.Add("@CheckFGPT", finalInspection.CheckFGPT);
                    objParameter.Add("@MetalDetectionDoc", finalInspection.MetalDetectionDoc);
                    objParameter.Add("@SealingSampleDoc", finalInspection.SealingSampleDoc);
                    objParameter.Add("@FabricApprovalDoc", finalInspection.FabricApprovalDoc);
                    break;
                case "Insp-Others":
                    sqlUpdCmd += $@"
update FinalInspection
 set    ProductionStatus = @ProductionStatus  ,
        OthersRemark = @OthersRemark    ,
        CFA= ''   ,
        InspectionResult= 'On-going'   ,
        ShipmentStatus= 'On Hold'   ,
        SubmitDate=null,
        EditName= @userID,
        EditDate= getdate()
where   ID = @FinalInspectionID
";
                    objParameter.Add("@FinalInspectionID", finalInspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionResult", finalInspection.InspectionResult);
                    objParameter.Add("@ShipmentStatus", finalInspection.ShipmentStatus);
                    objParameter.Add("@OthersRemark", finalInspection.OthersRemark);
                    objParameter.Add("@ProductionStatus", finalInspection.ProductionStatus);
                    break;
                default:
                    break;
            }

            ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }

        public void SubmitFinalInspection(FinalInspection finalInspection, string userID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string sqlUpdCmd = $@"
update FinalInspection
 set    ProductionStatus = @ProductionStatus  ,
        OthersRemark= @OthersRemark    ,
        CFA= @CFA   ,
        Clerk= @Clerk   ,
        InspectionResult= @InspectionResult   ,
        InspectionStep = 'Submit' ,
        ShipmentStatus= @ShipmentStatus   ,
        SubmitDate=getdate(),
        EditName= @userID,
        EditDate= getdate()
where   ID = @FinalInspectionID
";
            objParameter.Add("@FinalInspectionID", finalInspection.ID);
            objParameter.Add("@userID", userID);
            objParameter.Add("@InspectionResult", finalInspection.InspectionResult);
            objParameter.Add("@ShipmentStatus", finalInspection.ShipmentStatus);
            objParameter.Add("@ProductionStatus", finalInspection.ProductionStatus);
            objParameter.Add("@OthersRemark", finalInspection.OthersRemark);
            objParameter.Add("@CFA", finalInspection.CFA);
            objParameter.Add("@Clerk", finalInspection.Clerk);

            ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }
        public void UpdateGeneral(FinalInspectionGeneral General)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@FinalInspectionID", General.FinalInspectionID);

            objParameter.Add("@IsMaterialApproval", General.IsMaterialApproval);
            objParameter.Add("@IsSealingSample", General.IsSealingSample);
            objParameter.Add("@IsMetalDetection", General.IsMetalDetection);
            objParameter.Add("@IsFGWT", General.IsFGWT);
            objParameter.Add("@IsFGPT", General.IsFGPT);
            objParameter.Add("@IsTopSample", General.IsTopSample);
            objParameter.Add("@Is3rdPartyTestReport", General.Is3rdPartyTestReport);
            objParameter.Add("@IsPPSample", General.IsPPSample);
            objParameter.Add("@IsGBTestForChina", General.IsGBTestForChina);
            objParameter.Add("@IsCPSIAForYounthStytle", General.IsCPSIAForYounthStytle);
            objParameter.Add("@IsQRSSample", General.IsQRSSample);
            objParameter.Add("@IsFactoryDisclaimer", General.IsFactoryDisclaimer);
            objParameter.Add("@IsA01Compliance", General.IsA01Compliance);
            objParameter.Add("@IsCPSIACompliance", General.IsCPSIACompliance);
            objParameter.Add("@IsCustomerCountrySpecificCompliance", General.IsCustomerCountrySpecificCompliance);

            string sqlInsertFinalInspection_Moisture = @"
UPDATE dbo.FinalInspectionGeneral
   SET IsMaterialApproval = @IsMaterialApproval
      ,IsSealingSample = @IsSealingSample
      ,IsMetalDetection = @IsMetalDetection
      ,IsFGWT = @IsFGWT
      ,IsFGPT = @IsFGPT
      ,IsTopSample = @IsTopSample
      ,Is3rdPartyTestReport = @Is3rdPartyTestReport
      ,IsPPSample = @IsPPSample
      ,IsGBTestForChina = @IsGBTestForChina
      ,IsCPSIAForYounthStytle = @IsCPSIAForYounthStytle
      ,IsQRSSample = @IsQRSSample
      ,IsFactoryDisclaimer = @IsFactoryDisclaimer
      ,IsA01Compliance = @IsA01Compliance
      ,IsCPSIACompliance = @IsCPSIACompliance
      ,IsCustomerCountrySpecificCompliance = @IsCustomerCountrySpecificCompliance
WHERE FinalInspectionID = @FinalInspectionID
";
            ExecuteNonQuery(CommandType.Text, sqlInsertFinalInspection_Moisture, objParameter);
        }
        public void UpdateCheckList(FinalInspectionCheckList CheckList)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@FinalInspectionID", CheckList.FinalInspectionID);
            objParameter.Add("@IsCloseShade", CheckList.IsCloseShade);
            objParameter.Add("@IsHandfeel", CheckList.IsHandfeel);
            objParameter.Add("@IsAppearance", CheckList.IsAppearance);
            objParameter.Add("@IsPrintEmbDecorations", CheckList.IsPrintEmbDecorations);
            objParameter.Add("@IsEmbellishmentPrint", CheckList.IsEmbellishmentPrint);
            objParameter.Add("@IsEmbellishmentBonding", CheckList.IsEmbellishmentBonding);
            objParameter.Add("@IsEmbellishmentHT", CheckList.IsEmbellishmentHT);
            objParameter.Add("@IsEmbellishmentEMB", CheckList.IsEmbellishmentEMB);
            objParameter.Add("@IsFiberContent", CheckList.IsFiberContent);
            objParameter.Add("@IsCareInstructions", CheckList.IsCareInstructions);
            objParameter.Add("@IsDecorativeLabel", CheckList.IsDecorativeLabel);
            objParameter.Add("@IsAdicomLabel", CheckList.IsAdicomLabel);
            objParameter.Add("@IsCountryofOrigion", CheckList.IsCountryofOrigion);
            objParameter.Add("@IsSizeKey", CheckList.IsSizeKey);
            objParameter.Add("@Is8FlagLabel", CheckList.Is8FlagLabel);
            objParameter.Add("@IsAdditionalLabel", CheckList.IsAdditionalLabel);
            objParameter.Add("@IsIdLabel", CheckList.IsIdLabel);
            objParameter.Add("@IsMainLabel", CheckList.IsMainLabel);
            objParameter.Add("@IsSizeLabel", CheckList.IsSizeLabel);
            objParameter.Add("@IsCareContentLabel", CheckList.IsCareContentLabel);
            objParameter.Add("@IsBrandLabel", CheckList.IsBrandLabel);
            objParameter.Add("@IsBlueSignLabel", CheckList.IsBlueSignLabel);
            objParameter.Add("@IsLotLabel", CheckList.IsLotLabel);
            objParameter.Add("@IsSecurityLabel", CheckList.IsSecurityLabel);
            objParameter.Add("@IsSpecialLabel", CheckList.IsSpecialLabel);
            objParameter.Add("@IsVIDLabel", CheckList.IsVIDLabel);
            objParameter.Add("@IsCNC", CheckList.IsCNC);
            objParameter.Add("@IsWovenlabel", CheckList.IsWovenlabel);
            objParameter.Add("@IsTSize", CheckList.IsTSize);
            objParameter.Add("@IsCCLayout", CheckList.IsCCLayout);
            objParameter.Add("@IsShippingMark", CheckList.IsShippingMark);
            objParameter.Add("@IsPolytagMarking", CheckList.IsPolytagMarking);
            objParameter.Add("@IsColorSizeQty", CheckList.IsColorSizeQty);
            objParameter.Add("@IsHangtag", CheckList.IsHangtag);
            objParameter.Add("@IsJokerTag", CheckList.IsJokerTag);
            objParameter.Add("@IsWWMT", CheckList.IsWWMT);
            objParameter.Add("@IsChinaCIT", CheckList.IsChinaCIT);
            objParameter.Add("@IsPolybagSticker", CheckList.IsPolybagSticker);
            objParameter.Add("@IsUCCSticker", CheckList.IsUCCSticker);
            objParameter.Add("@IsPESheetMicropak", CheckList.IsPESheetMicropak);
            objParameter.Add("@IsAdditionalHantage", CheckList.IsAdditionalHantage);
            objParameter.Add("@IsUPCStickierHantage", CheckList.IsUPCStickierHantage);
            objParameter.Add("@IsGS1128Label", CheckList.IsGS1128Label);
            objParameter.Add("@IsSecuritytag", CheckList.IsSecuritytag);

            string sqlInsertFinalInspection_Moisture = @"
UPDATE dbo.FinalInspectionCheckList
   SET IsCloseShade = @IsCloseShade
      ,IsHandfeel = @IsHandfeel
      ,IsAppearance = @IsAppearance
      ,IsPrintEmbDecorations = @IsPrintEmbDecorations
      ,IsEmbellishmentPrint = @IsEmbellishmentPrint
      ,IsEmbellishmentBonding = @IsEmbellishmentBonding
      ,IsEmbellishmentHT = @IsEmbellishmentHT
      ,IsEmbellishmentEMB = @IsEmbellishmentEMB
      ,IsFiberContent = @IsFiberContent
      ,IsCareInstructions = @IsCareInstructions
      ,IsDecorativeLabel = @IsDecorativeLabel
      ,IsAdicomLabel = @IsAdicomLabel
      ,IsCountryofOrigion = @IsCountryofOrigion
      ,IsSizeKey = @IsSizeKey
      ,Is8FlagLabel = @Is8FlagLabel
      ,IsAdditionalLabel = @IsAdditionalLabel
      ,IsIdLabel = @IsIdLabel
      ,IsMainLabel = @IsMainLabel
      ,IsSizeLabel = @IsSizeLabel
      ,IsCareContentLabel = @IsCareContentLabel
      ,IsBrandLabel = @IsBrandLabel
      ,IsBlueSignLabel = @IsBlueSignLabel
      ,IsLotLabel = @IsLotLabel
      ,IsSecurityLabel = @IsSecurityLabel
      ,IsSpecialLabel = @IsSpecialLabel
      ,IsVIDLabel = @IsVIDLabel
      ,IsCNC = @IsCNC
      ,IsWovenlabel = @IsWovenlabel
      ,IsTSize = @IsTSize
      ,IsCCLayout = @IsCCLayout
      ,IsShippingMark = @IsShippingMark
      ,IsPolytagMarking = @IsPolytagMarking
      ,IsColorSizeQty = @IsColorSizeQty
      ,IsHangtag = @IsHangtag
      ,IsJokerTag = @IsJokerTag
      ,IsWWMT = @IsWWMT
      ,IsChinaCIT = @IsChinaCIT
      ,IsPolybagSticker = @IsPolybagSticker
      ,IsUCCSticker = @IsUCCSticker
      ,IsPESheetMicropak = @IsPESheetMicropak
      ,IsAdditionalHantage = @IsAdditionalHantage
      ,IsUPCStickierHantage = @IsUPCStickierHantage
      ,IsGS1128Label = @IsGS1128Label
      ,IsSecuritytag = @IsSecuritytag
WHERE FinalInspectionID = @FinalInspectionID
";
            ExecuteNonQuery(CommandType.Text, sqlInsertFinalInspection_Moisture, objParameter);
        }


        public IList<ImageRemark> GetFinalInspectionDetail(long FinalInspection_DetailUkey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@FinalInspection_DetailUkey", FinalInspection_DetailUkey }
            };

            string sqlGetData = @"
select  Image, Remark
from SciPMSFile_FinalInspection_DetailImage a with (nolock)
where   a.FinalInspection_DetailUkey = @FinalInspection_DetailUkey
";
            return ExecuteList<ImageRemark>(CommandType.Text, sqlGetData, objParameter);

        }

        public Dictionary<string, byte[]> GetFinalInspectionDefectImage(string FinalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@FinalInspectionID", DbType.String, FinalInspectionID }
            };

            string sqlGetData = @"
    select  [ImageName] =  CONCAT(fdi.ID, '_', isnull(fd.GarmentDefectCodeID, ''), '_', fdi.Ukey, '.png'), Image
    from SciPMSFile_FinalInspection_DetailImage fdi with (nolock)
    left join FinalInspection_Detail fd with (nolock) on fd.Ukey = fdi.FinalInspection_DetailUkey
    where   fdi.ID = @FinalInspectionID
    union all
    select  [ImageName] =  CONCAT(fdi.ID, '_', fdi.Ukey, '.png'), Image
    from SciPMSFile_FinalInspection_OtherImage fdi with (nolock)
    where   fdi.ID = @FinalInspectionID
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().ToDictionary(s => s["ImageName"].ToString(), s => (byte[])s["Image"]);
            }
            else
            {
                return new Dictionary<string, byte[]>();
            }
        }

        public Dictionary<string, byte[]> GetInlineInspectionDefectImage(string InspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@InspectionID", DbType.String, InspectionID }
            };

            string sqlGetData = @"
    select  [ImageName] =  CONCAT(fdi.InlineInspectionReportID, '_', fdi.Ukey, '.png'), fdi.Image
    from SciPMSFile_InlineInspection_DetailImage fdi with (nolock)
    where   fdi.InlineInspectionReportID = @InspectionID
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().ToDictionary(s => s["ImageName"].ToString(), s => (byte[])s["Image"]);
            }
            else
            {
                return new Dictionary<string, byte[]>();
            }
        }

        public Dictionary<string, byte[]> GetEndLineInspectionDefectImage(string InspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@InspectionID", DbType.String, InspectionID }
            };

            string sqlGetData = @"
    select  [ImageName] =  CONCAT(fdi.InspectionReportID, '_', fdi.Ukey, '.png'), fdi.Image
    from SciPMSFile_Inspection_DetailImage fdi with (nolock)
    where   fdi.InspectionReportID = @InspectionID
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().ToDictionary(s => s["ImageName"].ToString(), s => (byte[])s["Image"]);
            }
            else
            {
                return new Dictionary<string, byte[]>();
            }
        }

        public void UpdateFinalInspectionDetail(AddDefect addDefect, string UserID)
        {
            using (TransactionScope transaction = new TransactionScope())
            {
                SQLParameterCollection objParameter = new SQLParameterCollection() {
                    { "@FinalInspectionID", DbType.String, addDefect.FinalInspectionID },
                    { "@RejectQty", DbType.Int32, addDefect.RejectQty },
                };

                string sqlUpdFinalInspection = @"
update  FinalInspection
        set PassQty = SampleSize - @RejectQty,
            RejectQty = @RejectQty
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
                            { "@AreaCode", DbType.String, defectItem.AreaCode ?? string.Empty},
                            { "@Remark", DbType.String, defectItem.Remark ?? string.Empty},
                            { "@Ukey",  defectItem.Ukey },
                            { "@Qty", DbType.Int32, defectItem.Qty }
                        };

                    List<string> OperationList = string.IsNullOrEmpty(defectItem.Operation) ? new List<string>() : defectItem.Operation.Split(',').ToList();
                    List<string> OperatorList = string.IsNullOrEmpty(defectItem.Operator) ? new List<string>() : defectItem.Operator.Split(',').ToList();

                    if (defectItem.Ukey == -1)
                    {
                        sqlUpdateFinalInspectionDetail = @"
    DECLARE @FinalInspection_DetailKey table (Ukey bigint)

    insert into FinalInspection_Detail(ID, GarmentDefectTypeID, GarmentDefectCodeID ,AreaCode ,Remark ,Qty)
                OUTPUT INSERTED.Ukey into @FinalInspection_DetailKey
                values(@FinalInspectionID, @GarmentDefectTypeID, @GarmentDefectCodeID ,@AreaCode ,@Remark ,@Qty)

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
            set Qty = @Qty ,AreaCode = @AreaCode ,Remark = @Remark 
            where   Ukey = @Ukey
    end
    else
    begin
        delete  FinalInspection_Detail where   Ukey = @Ukey
    end
    ;

        delete  FinalInspection_Detail_Operation where   InspectionDetailUkey = @Ukey
";
                        ExecuteNonQuery(CommandType.Text, sqlUpdateFinalInspectionDetail, detailParameter);
                    }

                    if (defectItem.Qty > 0)
                    {
                        foreach (var DetailImage in defectItem.ListFinalInspectionDefectImage)
                        {
                            string sqlInsertFinalInspection_DetailImage = @"
SET XACT_ABORT ON
insert into SciPMSFile_FinalInspection_DetailImage(ID, FinalInspection_DetailUkey, Image ,Remark)
            values(@FinalInspectionID, @FinalInspection_DetailUkey, @Image ,@Remark)
";
                            SQLParameterCollection imgParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, addDefect.FinalInspectionID },
                            { "@FinalInspection_DetailUkey",  defectItem.Ukey },
                            { "@Image", DetailImage.Image == null ? System.Data.SqlTypes.SqlBinary.Null : DetailImage.Image},
                            { "@Remark",DbType.String, DetailImage.Remark ?? "" },
                        };

                            ExecuteNonQuery(CommandType.Text, sqlInsertFinalInspection_DetailImage, imgParameter);
                        }

                        int i = 0;
                        foreach (var strOperation in OperationList)
                        {
                            string strOperator = OperatorList[i];
                            string sqlInsertFinalInspection_DetailImage = @"
INSERT INTO FinalInspection_Detail_Operation
           (InspectionDetailUkey
           ,InlineOperation
           ,InlineOperator)
     VALUES
           (@FinalInspection_DetailUkey
           ,@Operation
           ,@Operator)
";
                            SQLParameterCollection imgParameter = new SQLParameterCollection() {
                            { "@FinalInspection_DetailUkey",  defectItem.Ukey },
                            { "@Operation",DbType.String, strOperation},
                            { "@Operator",DbType.String, strOperator},
                        };

                            ExecuteNonQuery(CommandType.Text, sqlInsertFinalInspection_DetailImage, imgParameter);
                            i++;
                        }
                    }
                }

                transaction.Complete();
                transaction.Dispose();
            }
        }


        public void UpdateFinalInspectionDefectDetail(AddDefect addDefect)
        {
            if (addDefect.FinalInspection_DefectDetails == null)
            {
                return;
            }
            using (TransactionScope transaction = new TransactionScope())
            {
                foreach (FinalInspection_DefectDetail finalInspection_DefectDetail in addDefect.FinalInspection_DefectDetails)
                {
                    string sqlFinalInspection_DefectDetail = string.Empty;
                    SQLParameterCollection detailParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, addDefect.FinalInspectionID },
                            { "@ProUkey", finalInspection_DefectDetail.ProUkey },
                            { "@DefectCategoryUkey", finalInspection_DefectDetail.DefectCategoryUkey },
                            { "@Qty", finalInspection_DefectDetail.Qty }
                        };

                    sqlFinalInspection_DefectDetail = $@"
if not exists( 
select 1 from FinalInspection_DefectDetail
where FinalInspectionID = @FinalInspectionID and ProUkey = @ProUkey and DefectCategoryUkey=@DefectCategoryUkey

)
begin
    INSERT INTO dbo.FinalInspection_DefectDetail
               (FinalInspectionID,ProUkey,DefectCategoryUkey,Qty)
         VALUES
               (@FinalInspectionID,@ProUkey,@DefectCategoryUkey,@Qty)
end
else
begin
    UPDATE FinalInspection_DefectDetail SET Qty = @Qty WHERE FinalInspectionID = @FinalInspectionID and ProUkey = @ProUkey AND DefectCategoryUkey=@DefectCategoryUkey
end
";


                    ExecuteNonQuery(CommandType.Text, sqlFinalInspection_DefectDetail, detailParameter);
                }


                transaction.Complete();
                transaction.Dispose();
            }
        }

        public IList<BACriteriaItem> GetBeautifulProductAuditForInspection(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetData = $@"
select ID, Description 
into #baseBACriteria
from  Production.dbo.DropDownList ddl  WITH(NOLOCK)
where Type = 'PMS_BACriteria'
order by Seq

select  [Ukey] = isnull(fn.Ukey, -1),
        [BACriteria] = bac.ID,
        [BACriteriaDesc] = bac.Description,
        [Qty] = isnull(fn.Qty, 0),		
		[RowIndex]=ROW_NUMBER() OVER(ORDER BY bac.ID) -1
		,HasImage = Cast(
			IIF(EXISTS(
				select 1 from SciPMSFile_FinalInspection_NonBACriteriaImage img 
				where img.FinalInspection_NonBACriteriaUkey = fn.Ukey
			),1,0)		
		as bit)
    from #baseBACriteria bac with (nolock)
    left join   FinalInspection_NonBACriteria fn WITH(NOLOCK) on    fn.ID = @finalInspectionID and
                                                            fn.BACriteria = bac.ID

DROP TABLE #baseBACriteria
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
                };

                string sqlUpdFinalInspection = @"
update  FinalInspection
        set BAQty = @BAQty
where   ID = @FinalInspectionID
";
                ExecuteNonQuery(CommandType.Text, sqlUpdFinalInspection, objParameter);

                foreach (BACriteriaItem criteriaItem in beautifulProductAudit.ListBACriteria)
                {
                    string sqlUpdateFinalInspectionDetail = string.Empty;
                    SQLParameterCollection detailParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, beautifulProductAudit.FinalInspectionID },
                            { "@BACriteria", DbType.String, criteriaItem.BACriteria },
                            { "@Ukey", criteriaItem.Ukey },
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
                        foreach (var baDetail in criteriaItem.ListBACriteriaImage)
                        {
                            string sqlInsertFinalInspection_NonBACriteriaImage = @"
SET XACT_ABORT ON

insert into SciPMSFile_FinalInspection_NonBACriteriaImage(ID, FinalInspection_NonBACriteriaUkey, Image ,Remark)
            values(@FinalInspectionID, @FinalInspection_NonBACriteriaUkey, @Image ,@Remark)
";
                            SQLParameterCollection imgParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, beautifulProductAudit.FinalInspectionID },
                            { "@FinalInspection_NonBACriteriaUkey",  criteriaItem.Ukey },
                            { "@Remark", DbType.String, baDetail.Remark ?? ""},
                            { "@Image", baDetail.Image}
                        };

                            ExecuteNonQuery(CommandType.Text, sqlInsertFinalInspection_NonBACriteriaImage, imgParameter);
                        }
                    }
                }

                transaction.Complete();
            }
        }

        public IList<ImageRemark> GetBA_DetailImage(long FinalInspection_NonBACriteriaUkey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@FinalInspection_NonBACriteriaUkey",  FinalInspection_NonBACriteriaUkey }
            };

            string sqlGetData = @"
select  Remark, Image
from SciPMSFile_FinalInspection_NonBACriteriaImage a with (nolock)
where   a.FinalInspection_NonBACriteriaUkey = @FinalInspection_NonBACriteriaUkey
";

            return ExecuteList<ImageRemark>(CommandType.Text, sqlGetData, objParameter);
        }

        public IList<CartonItem> GetMoistureListCartonItem(string finalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@finalInspectionID", DbType.String, finalInspectionID }
            };

            string sqlGetMoistureListCartonItem = @"
select distinct ID, Seq, Article
into    #Order_QtyShip_Detail
from    Production.dbo.Order_QtyShip_Detail WITH(NOLOCK)  
where   ID in (select OrderID from FinalInspection_Order with (nolock) where ID = @finalInspectionID)

select  [FinalInspection_OrderCartonUkey] = foc.Ukey,
        foc.OrderID,
        foc.PackinglistID,
        foc.CTNNo,
        [Article] = isnull(oqd.Article, '')
from    FinalInspection_OrderCarton foc with (nolock)
left join #Order_QtyShip_Detail oqd on oqd.ID = foc.OrderID and oqd.Seq = foc.Seq 
where foc.ID = @finalInspectionID

";
            return ExecuteList<CartonItem>(CommandType.Text, sqlGetMoistureListCartonItem, objParameter);

        }


        public IList<ViewMoistureResult> GetViewMoistureResult(string finalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@finalInspectionID", DbType.String, finalInspectionID }
            };

            string sqlGetViewMoistureResult = @"
declare @FinalInspection_CTNMoistureStandard numeric(5,2)
declare @BrandID as varchar(10) 

select @BrandID = BrandID
from FinalInspection a WITH(NOLOCK)
inner join FinalInspection_Order b WITH(NOLOCK) on a.ID=b.ID
inner join Production..Orders c WITH(NOLOCK) on c.ID=b.OrderID
where a.ID = @finalInspectionID


select @FinalInspection_CTNMoistureStandard = Standard
from FinalInspectionMoistureStandard
where Category = 'CTNMoisture' and BrandID = @BrandID

IF @FinalInspection_CTNMoistureStandard  IS NULL
BEGIN
	select @FinalInspection_CTNMoistureStandard = Standard
	from FinalInspectionMoistureStandard
	where Category = 'CTNMoisture' and BrandID = 'Default'
END

select *
into #EndlineMoisture
from EndlineMoisture WITH(NOLOCK)
where BrandID = @BrandID
and Junk = 0

select  fm.Ukey,
        fm.Article,
        fo.CTNNo,
        fm.Instrument,
        fm.Fabrication,
        [GarmentStandard] = ISNULL(em.Standard ,emDefault.Standard),
        fm.GarmentTop,
        fm.GarmentMiddle,
        fm.GarmentBottom,
        [CTNStandard] = @FinalInspection_CTNMoistureStandard,
        fm.CTNInside,
        fm.CTNOutside,
        Result = IIF(fm.Result = 'P','Pass','Fail'),
        fm.Action,
        fm.Remark
from    FinalInspection_Moisture fm with (nolock)
left join   FinalInspection_OrderCarton fo with (nolock) on fo.Ukey = fm.FinalInspection_OrderCartonUkey
left join   #EndlineMoisture em on em.Instrument = fm.Instrument and em.Fabrication = fm.Fabrication
left join   EndlineMoisture emDefault with (nolock) on emDefault.Instrument = fm.Instrument and emDefault.Fabrication = fm.Fabrication and emDefault.BrandID='' and emDefault.Junk = 0
where fm.ID = @finalInspectionID

drop table #EndlineMoisture
";
            return ExecuteList<ViewMoistureResult>(CommandType.Text, sqlGetViewMoistureResult, objParameter);
        }

        /// <summary>
        /// 根據品牌尋找EndlineMoisture設定
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="BrandID"></param>
        /// <returns></returns>
        public IList<EndlineMoisture> GetEndlineMoistureByBrand(string FinalInspectionID, string BrandID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@FinalInspectionID", DbType.String, FinalInspectionID }
            ,{ "@BrandID", DbType.String, BrandID }
            };

            string sqlGetEndlineMoisture = string.Empty;
            if (!string.IsNullOrEmpty(FinalInspectionID))
            {
                sqlGetEndlineMoisture = $@"
select distinct d.*
from FinalInspection a WITH(NOLOCK)
inner join FinalInspection_Order b WITH(NOLOCK) on a.ID=b.ID 
inner join Production..Orders c WITH(NOLOCK) on c.ID=b.OrderID
inner join EndlineMoisture d WITH(NOLOCK) on d.BrandID=c.BrandID and d.Junk = 0
where a.ID = @FinalInspectionID
";
            }
            else if (!string.IsNullOrEmpty(BrandID))
            {
                sqlGetEndlineMoisture = $@"
select *
from EndlineMoisture WITH(NOLOCK)
where BrandID = @BrandID
and Junk = 0
";
            }

            return ExecuteList<EndlineMoisture>(CommandType.Text, sqlGetEndlineMoisture, objParameter);
        }

        /// <summary>
        /// 取得EndlineMoisture的預設值
        /// </summary>
        /// <param name="FinalInspectionID"></param>
        /// <param name="BrandID"></param>
        /// <returns></returns>
        public IList<EndlineMoisture> GetEndlineMoistureDefault()
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlGetEndlineMoisture = @"
select  *
from    EndlineMoisture with (nolock)
where BrandID = ''
and Junk = 0
";
            return ExecuteList<EndlineMoisture>(CommandType.Text, sqlGetEndlineMoisture, objParameter);
        }

        /// <summary>
        /// 取得箱子濕度標準(FinalInspectionMoistureStandard)
        /// </summary>
        /// <param name="finalInspectionID"></param>
        /// <returns></returns>
        public List<FinalInspectionMoistureStandard> GetMoistureStandardSetting(string finalInspectionID)
        {

            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@FinalInspectionID", DbType.String, finalInspectionID }
            };

            string sqlGetEndlineMoisture = @"

select distinct d.*
into #BrandSetting
from FinalInspection a  with (nolock)
inner join FinalInspection_Order b with (nolock) on a.ID=b.ID
inner join Production..Orders c on c.ID=b.OrderID
inner join FinalInspectionMoistureStandard d with (nolock) on d.BrandID=c.BrandID
where a.ID = @FinalInspectionID

if not exists(
    select 1 from #BrandSetting
)
begin    
    select * 
    from FinalInspectionMoistureStandard with (nolock)
    where BrandID = 'Default'
end
else
begin
    select * from #BrandSetting
end

drop table #BrandSetting
";

            var tmp = ExecuteList<FinalInspectionMoistureStandard>(CommandType.Text, sqlGetEndlineMoisture, objParameter);
            return tmp.Any() ? tmp.ToList() : new List<FinalInspectionMoistureStandard>();
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
            objParameter.Add("@Result", DbType.String, moistureResult.Result);
            objParameter.Add("@Action", moistureResult.Action);
            objParameter.Add("@Remark", moistureResult.Remark ?? "");
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
,@Result
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

        public bool CheckMoistureExists(string finalInspectionID, string article, long finalInspection_OrderCartonUkey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string where = string.Empty;
            objParameter.Add("@FinalInspectionID", finalInspectionID);

            if (!string.IsNullOrEmpty(article))
            {
                where += " and Article = @article";
                objParameter.Add("@article", article);
            }

            if (finalInspection_OrderCartonUkey > 0)
            {
                where += " and FinalInspection_OrderCartonUkey = @finalInspection_OrderCartonUkey";
                objParameter.Add("@finalInspection_OrderCartonUkey", finalInspection_OrderCartonUkey);
            }


            string sqlCheckMoistureExists = $@"
declare @BrandID as varchar(10)

select  @BrandID = BrandID
from FinalInspection a with (nolock)
inner join FinalInspection_Order b with (nolock) on a.ID=b.ID
inner join Production..Orders c with (nolock) on c.ID=b.OrderID
where a.ID = @FinalInspectionID

----LLL 最多可以驗三次，其餘只能一次
if @BrandID  = 'LLL'
BEGIN
	select  [result] = IIF( COUNT(Ukey) >= 3,1,0)
	from    FinalInspection_Moisture with (nolock)
    where   ID = @FinalInspectionID {where}
END
ELSE
BEGIN
	select  [result] = 1 
	from    FinalInspection_Moisture with (nolock)
    where   ID = @FinalInspectionID {where}
END
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlCheckMoistureExists, objParameter);
            int ctn = Convert.ToInt32(ExecuteScalar(CommandType.Text, sqlCheckMoistureExists, objParameter));

            return ctn > 0;
        }

        public void InsertMeasurement(DatabaseObject.ViewModel.FinalInspection.ServiceMeasurement measurement, string userID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            DataTable dtDateTime = ExecuteDataTableByServiceConn(CommandType.Text, "select [DateTime] = getdate()", objParameter);


            using (TransactionScope transactionScope = new TransactionScope())
            {
                foreach (MeasurementItem measurementItem in measurement.ListMeasurementItem)
                {
                    objParameter = new SQLParameterCollection
                    {
                        { "@SizeUnit", measurement.SizeUnit },
                        { "@ID", measurement.FinalInspectionID },
                        { "@Article", measurement.SelectedArticle },
                        { "@SizeCode", measurement.SelectedSize },
                        { "@Location", measurement.SelectedProductType },
                        { "@Code", measurementItem.Code },
                        { "@SizeSpec", measurementItem.ResultSizeSpec },
                        { "@MeasurementUkey", measurementItem.MeasurementUkey },
                        { "@AddName", userID },
                        { "@AddDate", dtDateTime.Rows[0]["DateTime"] }
                    };

                    bool isFractional = false;
                    if (measurementItem.ResultSizeSpec != null && measurementItem.ResultSizeSpec.Contains("/") && !System.Text.RegularExpressions.Regex.IsMatch(measurementItem.ResultSizeSpec, @"[a-zA-Z]"))
                    {
                        isFractional = true;
                    }

                    string ins = "@SizeSpec";
                    if (isFractional)
                    {
                        ins = "(SELECT dbo.getFractional(@SizeSpec))";
                    }

                    string sqlInsertMeasurement = $@"
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
,{ins}
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

        public void DeleteMeasurement(DatabaseObject.ViewModel.FinalInspection.ServiceMeasurement measurement, DateTime AddDate)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ID", measurement.FinalInspectionID);
            objParameter.Add("@AddDate", AddDate);


            using (TransactionScope transactionScope = new TransactionScope())
            {
                string sql = @"delete from FinalInspection_Measurement where ID = @ID and CONVERT(varchar, AddDate,120) = @AddDate";
                ExecuteNonQuery(CommandType.Text, sql, objParameter);
                transactionScope.Complete();
            }
        }
        public void UpdateMeasurement(ServiceMeasurement model, string userID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add($@"@SizeUnit", model.SizeUnit);

            using (TransactionScope transactionScope = new TransactionScope())
            {
                StringBuilder stringBuilder = new StringBuilder();
                int idx = 0;
                foreach (var measuremen in model.ListMeasurementItem)
                {
                    objParameter.Add($@"@Ukey{idx}", measuremen.FinalInspection_MeasurementUkey);
                    objParameter.Add($@"@SizeSpec{idx}", measuremen.ResultSizeSpec);

                    bool isFractional = false;
                    if (measuremen.ResultSizeSpec != null && measuremen.ResultSizeSpec.Contains("/"))
                    {
                        isFractional = true;
                    }

                    string ins = $@"@SizeSpec{idx}";
                    if (isFractional)
                    {
                        ins = $@"(SELECT dbo.getFractional(@SizeSpec{idx}))";
                    }

                    string sql = $@"
UPDATE FinalInspection_Measurement 
SET EditDate=GETDATE()
,EditName = '{userID}'
,SizeSpec = {ins}
where Ukey = @Ukey{idx}";
                    stringBuilder.Append(sql);
                    idx++;
                }
                ExecuteNonQuery(CommandType.Text, stringBuilder.ToString(), objParameter);
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
Order BY Article,SizeCode,Location

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

declare @CustPONO varchar(30)

select  @CustPONO = CustPONO
from    FinalInspection with (nolock)
where   ID = @finalInspectionID


select  StyleUkey = Ukey,SizeUnit
INTO #Style_Size
from    Production.dbo.Style WITH(NOLOCK)
where   Ukey IN (
	select StyleUkey 
	from Production.dbo.Orders o WITH(NOLOCK)
	INNER JOIN FinalInspection_Order fo WITH(NOLOCK) on fo.OrderID=o.ID
	INNER JOIN FinalInspection f WITH(NOLOCK) on f.ID=fo.ID
	where f.ID =  @finalInspectionID
    ----因為CustPONO會有空值的情況，所以不可以使用CustPONO去取Style，改成直接抓SP#
)

select  SizeSpec,        MeasurementUkey,        AddDate    ,FinalInspection_MeasurementUkey = Ukey
into    #tmp_Inspection_Measurement
from    FinalInspection_Measurement WITH(NOLOCK)
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
	,[diff]= max(dbo.calculateSizeSpec(m.SizeSpec,im.SizeSpec, ss.SizeUnit))
	,im.AddDate
	,[HeadSizeCode] = FORMAT(im.AddDate,'yyyy/MM/dd HH:mm:ss')
	,im.FinalInspection_MeasurementUkey
into #tmp 
from Measurement m with(nolock)
INNER JOIN #Style_Size ss WITH(NOLOCK) ON m.StyleUkey = ss.StyleUkey 
left join #tmp_Inspection_Measurement im WITH(NOLOCK) on im.MeasurementUkey = m.Ukey 
LEFT JOIN [ManufacturingExecution].[dbo].[MeasurementTranslate] b WITH(NOLOCK) ON  m.MeasurementTranslateUkey = b.UKey
where  m.SizeCode = @size and m.junk = 0
AND (m.SizeSpec NOT LIKE '%!%' AND m.SizeSpec NOT LIKE '%@%' AND m.SizeSpec NOT LIKE '%#%' 
AND m.SizeSpec NOT LIKE '%$%'  AND m.SizeSpec NOT LIKE '%^%'  AND m.SizeSpec NOT LIKE '%&%' 
AND m.SizeSpec NOT LIKE '%*%' AND m.SizeSpec NOT LIKE '%=%' AND m.SizeSpec NOT LIKE '%-%' 
AND m.SizeSpec NOT LIKE '%(%' AND m.SizeSpec NOT LIKE '%)%')
group by m.Ukey,iif(isnull(b.DescEN,'') = '',m.Description,b.DescEN),m.Tol1,m.Tol2,m.Code,m.SizeCode,m.SizeSpec,im.SizeSpec,im.AddDate,im.FinalInspection_MeasurementUkey

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
		,Max(case when HeadSizeCode ='''+@HeadSizeCode+''' and SizeCode ='''+@mSizeCode+''' then FinalInspection_MeasurementUkey end) as FinalInspection_MeasurementUkey'+@r_id+'
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


drop table #tmp,#Style_Size

";

            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);
            return dt;
        }


        public IList<OtherImage> GetOthersImageList(string finalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@finalInspectionID", finalInspectionID);

            string sqlGetEndlineMoisture = @"
select  Ukey
    , ID
    , Image
    , Remark
    ,[RowIndex]=ROW_NUMBER() OVER(ORDER BY Ukey) -1
from SciPMSFile_FinalInspection_OtherImage with (nolock)
where   ID = @finalInspectionID

";
            return ExecuteList<OtherImage>(CommandType.Text, sqlGetEndlineMoisture, objParameter);
        }

        public void UpdateFinalInspection_OtherImage(string finalInspectionID, List<OtherImage> images)
        {
            foreach (var imageObj in images)
            {
                string sqlFinalInspection_OtherImage = @"
SET XACT_ABORT ON

    insert into SciPMSFile_FinalInspection_OtherImage(ID, Image, Remark)
                values(@FinalInspectionID, @Image, @Remark)
";
                SQLParameterCollection imgParameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, finalInspectionID },
                            { "@Remark", DbType.String, imageObj.Remark ?? ""},
                            { "@Image", imageObj.Image}
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
declare @BuyerDelivery varchar(10)
declare @DiffDay int

select  @StyleID = StyleID,
        @SeasonID = SeasonID,
        @BrandID = BrandID,
        @BuyerDelivery = Format(BuyerDelivery, 'yyyy/MM/dd'),
		@DiffDay = (SELECT CAST( DATEDIFF(DAY ,GETDATE() ,BuyerDelivery) as int ))
from    Production.dbo.Orders with (nolock)
where   ID IN (
    select ID
    from Production.dbo.Orders WITH(NOLOCK)
    where CustPONO = (select CustPONO from FinalInspection with (nolock) where ID = @FinalInspectionID )
)

select  f.CustPONO,
        f.InspectionResult,
        f.FactoryID,
        f.InspectionStage,
        [SP] = (SELECT Stuff((select concat( ',',OrderID) 
                                from  FinalInspection_Order fo with (nolock) 
                                where fo.ID = f.ID
                                FOR XML PATH('')),1,1,'') ),
        [StyleID] = @StyleID,
        [SeasonID] = @SeasonID,
        [BrandID] = @BrandID,
        [BuyerDelivery] = @BuyerDelivery,
		[AiComment] = Ai.Comment,
        [CFA] = (select name from pass1 with (nolock) where ID = f.CFA),
        [SubmitDate] = Format(f.SubmitDate, 'yyyy/MM/dd'),
        [AuditDate] = Format(f.AuditDate, 'yyyy/MM/dd')
from FinalInspection f with (nolock)
OUTER APPLY(
	select ad.Comment
	from AIComment a
	inner join AIComment_Detail ad on a.Ukey = ad.AICommentUkey
	where a.FunctionName = 'CFA' + f.InspectionStage
	AND ad.[Day]= CASE WHEN @DiffDay <= 0 THEN 0
					   WHEN @DiffDay > 8 THEN 9
					   ELSE @DiffDay
					END
)Ai
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
declare @Customize4 varchar(50)
declare @Dest varchar(40)
declare @spQty int

select @spQty = isnull(sum(Qty), 0)
from Production.dbo.Orders with (nolock)
where   ID in (select OrderID from FinalInspection_Order with (nolock) where ID = @FinalInspectionID)

select  @StyleID = o.StyleID,
        @SeasonID = o.SeasonID,
        @BrandID = o.BrandID,
        @Customize4 = o.Customize4,
		@Dest = IIF(o.VasShas=1,o.Dest + '-' + c.Alias, '')
from    Production.dbo.Orders o with (nolock)
left join Production.dbo.Country c on o.Dest=c.ID
where   o.ID IN (
    select ID
    from Production.dbo.Orders WITH(NOLOCK)
    where CustPONO = (select CustPONO from FinalInspection with (nolock) where ID = @FinalInspectionID ) AND CustPONO != ''
	UNION
	select TOP 1 ID = OrderID 
	from FinalInspection_Order with (nolock) where ID = @FinalInspectionID
)

select  [SP] = (SELECT Stuff((select concat( ',',OrderID) 
                                from  FinalInspection_Order fo with (nolock) 
                                where fo.ID = @FinalInspectionID
                                FOR XML PATH('')),1,1,'') ),
        [StyleID] = @StyleID,
        [SeasonID] = @SeasonID,
        [BrandID] = @BrandID,
        [Customize4] = @Customize4,
        [CFA] = ISNULL((select name from pass1 with (nolock) where ID = f.CFA),(select name from Production..pass1 with (nolock) where ID = f.CFA)),
        [TotalSPQty] = @spQty,
		Dest = @Dest
from    FinalInspection f with (nolock)
where ID = @FinalInspectionID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, parameter);
        }

        public IList<QueryFinalInspection> GetFinalinspectionQueryList(QueryFinalInspection_ViewModel request)
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

            if (!string.IsNullOrEmpty(request.CustPONO))
            {
                whereOrder += @" and ID IN (
select ID
from Production.dbo.Orders WITH(NOLOCK)
where CustPONO = @CustPONO
)
";
                whereFinalInspection += " and exists (select 1 from SciProduction_Orders o WITH(NOLOCK) where fo.OrderID = o.ID and o.CustPONO = @CustPONO) ";
                parameter.Add("@CustPONO", request.CustPONO);
            }

            if (!string.IsNullOrEmpty(request.SP))
            {
                whereOrder += @" and ID = @SP";
                whereFinalInspection += @" and fo.OrderID = @SP";
                parameter.Add("@SP", request.SP);
            }

            if (request.AuditDateStart != null)
            {
                whereFinalInspection += @" and AuditDate >= @AuditDateStart";
                parameter.Add("@AuditDateStart", request.AuditDateStart);
            }

            if (request.AuditDateEnd != null)
            {
                whereFinalInspection += @" and AuditDate <= @AuditDateEnd";
                parameter.Add("@AuditDateEnd", request.AuditDateEnd);
            }

            if (request.SubmitDateStart != null)
            {
                whereFinalInspection += @" and SubmitDate >= @SubmitDateStart";
                parameter.Add("@SubmitDateStart", request.SubmitDateStart);
            }

            if (request.SubmitDateEnd != null)
            {
                whereFinalInspection += @" and SubmitDate <= @SubmitDateEnd";
                parameter.Add("@SubmitDateEnd", request.SubmitDateEnd);
            }

            if (!string.IsNullOrEmpty(request.StyleID))
            {
                whereOrder += @" and StyleID = @StyleID";
                parameter.Add("@StyleID", request.StyleID);
            }

            if (request.ExcludeJunk)
            {
                whereFinalInspection += @" and f.InspectionResult <> 'Junk' ";
            }

            string sqlGetData = $@"

select  ID,
        StyleID,
        SeasonID,
        BrandID,
        Customize4,
        Qty
into    #tmpOrders
from    Production.dbo.Orders with (nolock)
where   1 = 1 {whereOrder}

select  ID, Article
into    #tmpOrderArticle
from    Production.dbo.Order_Article with (nolock)
where   ID in (select ID from #tmpOrders)

select  [FinalInspectionID] = f.ID,
        [SP] = fo.OrderID,
        f.CustPONO,
        [AuditDate] = format(f.AuditDate, 'yyyy/MM/dd'),
        [SPQty] = cast(o.Qty as varchar),
        [StyleID] = o.StyleID,
        [Season] = o.SeasonID,
        [BrandID] = o.BrandID,
        o.Customize4,
        [Article] = (SELECT Stuff((select concat( ',',Article)   from #tmpOrderArticle where ID = fo.OrderID FOR XML PATH('')),1,1,'') ),
        [InspectionTimes] = cast(f.InspectionTimes as varchar),
        f.InspectionStage,
        f.InspectionResult,
		[IsTransferToPMS] = c.val,
		[IsTransferToPivot88] = iif(f.IsExportToP88 = 1, 'Y', 'N'),
        [SampleSize] = cast(f.SampleSize as varchar),
        [SubmitDate] = format(f.SubmitDate, 'yyyy/MM/dd')   
        ,f.ReInspection
from FinalInspection f with (nolock)
inner join FinalInspection_Order fo with (nolock) on fo.ID = f.ID
inner join #tmpOrders o on fo.OrderID = o.ID
outer apply(
	select val = iif(exists(
		select ID 
		from SciProduction_CFAInspectionRecord c with (nolock) 
		where c.ID = f.ID
	), 'Y', 'N')
)c
where   1 = 1 {whereFinalInspection}
";

            return ExecuteList<QueryFinalInspection>(CommandType.Text, sqlGetData, parameter);
        }
        public IList<QueryFinalInspection> GetFinalinspectionQueryList_Default(QueryFinalInspection_ViewModel request)
        {
            SQLParameterCollection parameter = new SQLParameterCollection();

            string whereOrder = string.Empty;
            string whereFinalInspection = string.Empty;

            if (request.ExcludeJunk)
            {
                whereFinalInspection = " where f.InspectionResult <> 'Junk' ";
            }

            string sqlGetData = $@"
--預設抓兩百
select distinct　top 200  f.ID,fo.OrderID,f.AddDate
into #default
from FinalInspection f with (nolock)
inner join  FinalInspection_Order fo with (nolock) on fo.ID = f.ID
Order by f.AddDate desc


select  ID, Article
into    #tmpOrderArticle
from    SciProduction_Order_Article with (nolock)
where   ID in (select OrderID from #default)


select top 200 [FinalInspectionID] = f.ID,
        [SP] = fo.OrderID,
        f.CustPONO,
        [AuditDate] = format(f.AuditDate, 'yyyy/MM/dd'),
        [SPQty] = cast(o.Qty as varchar),
        [StyleID] = o.StyleID,
        [Season] = o.SeasonID,
        [BrandID] = o.BrandID,
        o.Customize4,
        [Article] = (SELECT Stuff((select concat( ',',Article)   from #tmpOrderArticle where ID = fo.OrderID FOR XML PATH('')),1,1,'') ),
        [InspectionTimes] = cast(f.InspectionTimes as varchar),
        f.InspectionStage,
        f.InspectionResult,
        f.AddDate,
		[IsTransferToPMS] = c.val,
		[IsTransferToPivot88] = iif(f.IsExportToP88 = 1, 'Y', 'N'),
        [SampleSize] = cast(f.SampleSize as varchar),
        [SubmitDate] = format(f.SubmitDate, 'yyyy/MM/dd')  
        ,f.ReInspection
from FinalInspection f with (nolock)
inner join #default fo with (nolock) on fo.ID = f.ID
inner join Production.dbo.Orders o with(nolock) on o.ID = fo.OrderID
outer apply(
	select val = iif(exists(
		select ID 
		from SciProduction_CFAInspectionRecord c with(nolock)
		where c.ID = f.ID
	), 'Y', 'N')
)c 
{whereFinalInspection}
order by f.AddDate DESC

drop table #default ,#tmpOrderArticle
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

        public IList<SelectOrderShipSeq> GetListShipModeSeq(string finalInspectionID)
        {
            SQLParameterCollection parameter = new SQLParameterCollection() {
                            { "@FinalInspectionID", DbType.String, finalInspectionID }
                        };

            string sqlGetData = @"
select  OrderID
        ,Seq
        ,ShipmodeID
from    FinalInspection_Order_QtyShip with (nolock)
where   ID = @FinalInspectionID
";

            return ExecuteList<SelectOrderShipSeq>(CommandType.Text, sqlGetData, parameter);
        }

        #endregion

        public IList<DatabaseObject.ProductionDB.System> GetSystem()
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT * FROM [System] WITH(NOLOCK)" + Environment.NewLine);

            return ExecuteList<DatabaseObject.ProductionDB.System>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public DataSet GetEndInlinePivot88(string ID, string inspectionType)
        {
            SQLParameterCollection parameter = new SQLParameterCollection() {
                            { "@InspectionID", DbType.String, ID }
                        };

            string inspectionReportTable = inspectionType == "InlineInspection" ? "InlineInspectionReport" : "InspectionReport";
            string inspectionTable = inspectionType == "InlineInspection" ? "InlineInspection" : "Inspection";
            string dynamicCol = string.Empty;

            if (inspectionType == "InlineInspection")
            {
                dynamicCol = @"
[SewerID] = SUBSTRING(r.Operator,1, 255),
[Station] = SUBSTRING(r.Station,1, 255),
[Line] = SUBSTRING(r.Line,1, 255),
[Operation] = SUBSTRING(r.Operation,1, 255),
[Size] = '',
[WFT]='',
[EndlineCGradeQty] = Breakdown.RejectQty - Breakdown.FixQty
";
            }
            else
            {
                dynamicCol = @"
[SewerID] = '',
[Station] = '',
[Line] = '',
[Operation] = '',
[Size] = (SELECT val =  Stuff((select distinct concat( ',', isb.SizeCode) 
                from InspectionReport_Breakdown isb with (nolock)
                where isb.InspectionReportID = r.ID FOR XML PATH('')),1,1,'')),
[WFT]= r.WFT,
r.EndlineCGradeQty
";
            }


            string sqlGetData = $@"
declare @ID varchar(13) = @InspectionID

select  [FirstInspectionDate] = format(dateadd(hour, -system.UTCOffsetTimeZone, r.FirstInspectionDate), 'yyyy-MM-ddTHH:mm:ss'),
        [DefectQty] = Breakdown_Detail.Qty,
        Breakdown.PassQty,
        Breakdown.RejectQty,
        Breakdown.FixQty,
        [username] = (select Pivot88UserName from pass1 where id = r.QC),
        [LastinspectionDate] = format(dateadd(hour, -system.UTCOffsetTimeZone, r.LastinspectionDate), 'yyyy-MM-ddTHH:mm:ss'),
        [POQty] = orderInfo.Qty,
        [Customize5] = OrderJoinString.Customize5,
        [Customize4] = OrderJoinString.Customize4,
        [BuyerDelivery] = format(orderInfo.BuyerDelivery, 'yyyy-MM-ddTHH:mm:ss'),
        [Color] = Color.val,
        r.Shift,
        r.FactoryID,
        [InspectionMinutes] = Round(DATEDIFF(SECOND, r.FirstInspectionDate, r.LastinspectionDate) / 60.0, 0),
        r.CustPONO,
        orderInfo.CustCDID,
        OrderJoinString.SewLine,
        [Dest] = OrderJoinString.Dest,
        {dynamicCol}
from {inspectionReportTable} r with (nolock)
cross join system
outer apply(select  [RejectQty] = isnull(sum(RejectQty), 0),
                    [PassQty] = isnull(sum(PassQty), 0),
                    [FixQty] = isnull(sum(FixQty), 0)
            from {inspectionReportTable}_Breakdown with (nolock) where {inspectionReportTable}ID = r.ID) Breakdown
outer apply(select [Qty] = isnull(sum(Qty), 0) 
            from {inspectionReportTable}_Breakdown_Detail with (nolock) where {inspectionReportTable}ID = r.ID) Breakdown_Detail
outer apply(select  [Qty] = Sum(o.Qty),
                    [BuyerDelivery] = max(o.BuyerDelivery),  
                    [CustCDID] = max(o.CustCDID)
                    from Production.dbo.Orders o with (nolock) where o.CustPONO = r.CustPONO
            ) orderInfo
outer apply (
    SELECT  Customize4 = stuff((select distinct concat(',', o.Customize4) 
                         from Production.dbo.Orders o with (nolock) 
                         where o.CustPONO = r.CustPONO
                         FOR XML PATH('')), 1, 1, '')
            ,Customize5 = stuff((select distinct concat(',', o.Customize5) 
                        from Production.dbo.Orders o with (nolock)
                        where o.CustPONO = r.CustPONO
                        FOR XML PATH('')), 1, 1, '')
		    ,SewLine = (
			    select TOP 1 i.Line
                from Production.dbo.Orders o with (nolock) 
			    inner join {inspectionTable} i with (nolock) on i.OrderID = o.ID
                where o.CustPONO = r.CustPONO
		)
            ,Dest = stuff((select distinct concat(',', o.Dest) 
                        from Production.dbo.Orders o with (nolock)
                        where o.CustPONO = r.CustPONO
                        FOR XML PATH('')), 1, 1, '')
) OrderJoinString
outer apply(SELECT val =  Stuff((select distinct concat( ',', oc.ColorID) 
                from Production.dbo.Order_ColorCombo oc 
                where oc.ID in (select POID from Production.dbo.Orders where id in (select OrderID from {inspectionReportTable}_Breakdown where {inspectionReportTable}ID = r.ID))
            FOR XML PATH('')),1,1,'') ) Color
where r.ID = @ID

--取得Sku資料
if('{inspectionType}' = 'InlineInspection')
begin
    select  oq.Article,
            oq.SizeCode,
            [Qty] = sum(oq.Qty)
    into #tmpArticleSize
    from    Production.dbo.Order_Qty oq with (nolock)
    where   exists( select 1 
                    from InlineInspectionReport_Breakdown irb with (nolock) 
                    where   irb.InlineInspectionReportID = @ID and
                            irb.OrderID = oq.ID )
    group by oq.Article, oq.SizeCode    
    
    select  Article,
            SizeCode,
            [SizeRatio] = Qty * 1.0 / sum(Qty) over (partition by Article),
            [Seq] = ROW_NUMBER() OVER (PARTITION BY Article ORDER BY Qty desc)
    into #tmpSizeRatio
    from    #tmpArticleSize

    select  ta.Article,
            [Qty] = sum(isnull(irb.PassQty, 0) + isnull(irb.RejectQty, 0))
    into #inlineArticle
    from    (select distinct Article from #tmpArticleSize) ta
    left join InlineInspectionReport_Breakdown irb with (nolock) on irb.Article = ta.Article and irb.InlineInspectionReportID = @ID
    group by ta.Article

    ;WITH CTE_Letters AS (
        select	Article,
			    a.SizeCode,
			    [ShipQty] = case when ShipQty = 0 then 0
                                 when isLast = 0 then TotalQty - LAG(GrandQty, 1, 0) OVER (PARTITION BY Article ORDER BY Seq)
						    else ShipQty end,
		        (ROW_NUMBER() OVER (ORDER BY 
			        CAST(SUBSTRING( a.SizeCode, 1, PATINDEX('%[^0-9]%', a.SizeCode + 'Z') - 1) AS INT), 
			        SUBSTRING(a.SizeCode, PATINDEX('%[^0-9]%', a.SizeCode + 'Z'), LEN(a.SizeCode))
		        ) * 10) AS SeqNumber
	    from (	select  ia.Article,
			            sr.SizeCode,
			            [ShipQty] = FLOOR(ia.Qty * sr.SizeRatio),
					    [GrandQty] = sum(FLOOR(ia.Qty * sr.SizeRatio))  OVER (PARTITION BY sr.Article ORDER BY sr.Seq),
					    sr.Seq,
					    [isLast] = LEAD(sr.SizeRatio, 1, 0) OVER (PARTITION BY sr.Article ORDER BY sr.Seq),
					    [TotalQty] = ia.Qty
			    from    #inlineArticle ia
			    inner join  #tmpSizeRatio sr on sr.Article = ia.Article) a
        where   PATINDEX('%[^0-9]%', a.SizeCode) > 0 -- 包含字母的SizeCode		
    ),
    CTE_Numbers AS (
        select	Article,
			    a.SizeCode,
			    [ShipQty] = case when ShipQty = 0 then 0
                                 when isLast = 0 then TotalQty - LAG(GrandQty, 1, 0) OVER (PARTITION BY Article ORDER BY Seq)
						    else ShipQty end,
		        ((ROW_NUMBER() OVER (ORDER BY 
			        CAST(a.SizeCode AS INT)
		        ) + (SELECT COUNT(*) FROM CTE_Letters)) * 10) AS SeqNumber
	    from (	select  ia.Article,
			            sr.SizeCode,
			            [ShipQty] = FLOOR(ia.Qty * sr.SizeRatio),
					    [GrandQty] = sum(FLOOR(ia.Qty * sr.SizeRatio))  OVER (PARTITION BY sr.Article ORDER BY sr.Seq),
					    sr.Seq,
					    [isLast] = LEAD(sr.SizeRatio, 1, 0) OVER (PARTITION BY sr.Article ORDER BY sr.Seq),
					    [TotalQty] = ia.Qty
			    from    #inlineArticle ia
			    inner join  #tmpSizeRatio sr on sr.Article = ia.Article) a
        where   PATINDEX('%[^0-9]%', a.SizeCode) = 0 -- 純數字的SizeCode	    
    )
    SELECT * FROM CTE_Letters
    UNION ALL
    SELECT * FROM CTE_Numbers
    ORDER BY SeqNumber;

    drop table #tmpArticleSize, #tmpSizeRatio, #inlineArticle
end
else
begin
    ;WITH CTE_Letters AS (
        select  oq.Article,
                oq.SizeCode,
                [ShipQty] = sum(isnull(irb.PassQty, 0) + isnull(irb.RejectQty, 0)),
		        (ROW_NUMBER() OVER (ORDER BY 
			        CAST(SUBSTRING( oq.SizeCode, 1, PATINDEX('%[^0-9]%', oq.SizeCode + 'Z') - 1) AS INT), 
			        SUBSTRING( oq.SizeCode, PATINDEX('%[^0-9]%', oq.SizeCode + 'Z'), LEN(oq.SizeCode))
		        ) * 10) AS SeqNumber
        from    Production.dbo.Order_Qty oq with (nolock)
        left join InspectionReport_Breakdown irb with (nolock) on   irb.InspectionReportID = @ID and 
                                                                    irb.OrderID = oq.ID and 
                                                                    irb.Article = oq.Article and
                                                                    irb.SizeCode = oq.SizeCode
        where   oq.ID in (select OrderID from InspectionReport_Breakdown where InspectionReportID = @ID)
		AND PATINDEX('%[^0-9]%', oq.SizeCode) > 0 -- 包含字母的SizeCode		
        group by oq.Article, oq.SizeCode
    ),
    CTE_Numbers AS (
        select  oq.Article,
                oq.SizeCode,
                [ShipQty] = sum(isnull(irb.PassQty, 0) + isnull(irb.RejectQty, 0)),
		        ((ROW_NUMBER() OVER (ORDER BY 
			        CAST(oq.SizeCode AS INT)
		        ) + (SELECT COUNT(*) FROM CTE_Letters)) * 10) AS SeqNumber
        from    Production.dbo.Order_Qty oq with (nolock)
        left join InspectionReport_Breakdown irb with (nolock) on   irb.InspectionReportID = @ID and 
                                                                    irb.OrderID = oq.ID and 
                                                                    irb.Article = oq.Article and
                                                                    irb.SizeCode = oq.SizeCode
        where   oq.ID in (select OrderID from InspectionReport_Breakdown where InspectionReportID = @ID)
		AND PATINDEX('%[^0-9]%', oq.SizeCode) = 0 -- 純數字的SizeCode	
        group by oq.Article, oq.SizeCode
    )
    SELECT * FROM CTE_Letters
    UNION ALL
    SELECT * FROM CTE_Numbers
    ORDER BY SeqNumber;
end

select	s.StyleName,
		o.FactoryID,
		o.BrandID,
		s.CDCodeNew
into #tmpStyleInfo
from Production.dbo.Orders o with (nolock)
inner join Production.dbo.Style s with (nolock) on s.Ukey = o.StyleUkey
where   o.id in (select OrderID from {inspectionReportTable}_Breakdown where {inspectionReportTable}ID = @ID)

declare @AdidasSAPERPCode varchar(3)

select top 1 @AdidasSAPERPCode = SUBSTRING(BrandAreaCode, 1, 3)
from SciProduction_Factory_BrandDefinition fb with (nolock)
where   exists (select 1 from #tmpStyleInfo where fb.ID = FactoryID and fb.BrandID = BrandID and (fb.CDCodeID = CDCodeNew or fb.CDCodeID = ''))
order by fb.CDCodeID desc

--style info
SELECT	[Style] =  Stuff((select concat( ';',StyleName)   from #tmpStyleInfo FOR XML PATH('')),1,1,''),
		[BrandAreaCode] = @AdidasSAPERPCode,
		[BrandAreaID] = (select Name from Production.dbo.DropDownList with (nolock) where Type = 'AdidasSAPERPCode' and ID = @AdidasSAPERPCode)


select  [label] = isnull(gdt.Description, ''),
        [subsection] = isnull(gdc.Description, ''),
        [code] = isnull(gdc.Pivot88DefectCodeID, ''),
        [CriticalQty] = sum(iif(isnull(gdc.IsCriticalDefect, 0) = 1, ibd.Qty, 0)),
        [MajorQty] = sum(iif(isnull(gdc.IsCriticalDefect, 0) = 0, ibd.Qty, 0)),
        ibd.GarmentDefectCodeID
from {inspectionReportTable}_Breakdown_Detail ibd with (nolock)
left join SciProduction_GarmentDefectCode gdc with (nolock) on gdc.ID = ibd.GarmentDefectCodeID
left join SciProduction_GarmentDefectType gdt with (nolock) on gdt.ID = gdc.GarmentDefectTypeID
where ibd.{inspectionReportTable}ID = @ID
group by isnull(gdt.Description, ''), isnull(gdc.Description, ''), isnull(gdc.Pivot88DefectCodeID, ''), ibd.GarmentDefectCodeID

select  [title] = CONCAT(img.{inspectionReportTable}ID, '_', img.Ukey),
        [full_filename] = CONCAT(img.{inspectionReportTable}ID, '_', img.Ukey, '.png'),
        [number] = ROW_NUMBER() OVER (PARTITION BY insd.GarmentDefectCodeID ORDER BY img.Ukey),
        insd.GarmentDefectCodeID
from PMSFile.dbo.{inspectionTable}_DetailImage img with (nolock)
inner join {inspectionTable}_Detail insd with (nolock) on img.{inspectionTable}_DetailUkey = insd.Ukey
where img.{inspectionReportTable}ID = @ID

";
            return ExecuteDataSet(CommandType.Text, sqlGetData, parameter);
        }

        public DataSet GetPivot88(string ID)
        {
            SQLParameterCollection parameter = new SQLParameterCollection() {
                            { "@InspectionID", DbType.String, ID }
                        };

            string sqlGetData = @"
declare @ID varchar(13) = @InspectionID
declare @RandomHours FLOAT = (RAND() * (2.2 - 1.2)) + 1.2; -- 生成隨機小時數，範圍 1.2~2.2

Select	
    [AuditDate] = format(f.AuditDate, 'yyyy-MM-ddTHH:mm:ss'),
    f.RejectQty,
    f.InspectionResult,
    [DefectQty] = (select isnull(sum(Qty), 0) from FinalInspection_Detail with (nolock) where ID = f.ID),
    [AvailableQty] = (select sum(AvailableQty) from FinalInspection_Order with (nolock) where ID = f.ID),
    f.SampleSize,
    [InspectionLevel] = 
        case 
            when f.AcceptableQualityLevelsUkey = -1 then '100%inspection'
            when OrderInfo.IsDestJP = 1 or AQL.InspectionLevels = 2 then 'Japan orders (AQL 1.0, Level II)'
            else 'Regular orders (AQL 1.0, Level I)' 
        end,
    [InspectionResultID] = iif(f.InspectionResult = 'Pass', 1, 2),
    [InspectionStatusID] = iif(f.InspectionResult = 'Pass', 3, 7),
    f.SubmitDate,
    [InspectionMinutes] = Round( @RandomHours * 3600 / 60.0, 0),  -- 時間差
    [CFA] = isnull((select Pivot88UserName from quality_pass1 with (nolock) where ID = f.CFA), ''),
    OrderInfo.POQty,
    [ETD_ETA] = format(OrderInfo.ETD_ETA, 'yyyy-MM-ddTHH:mm:ss'),
    [CustPONO] = OrderJoinString.CustPONO,
    [Customize5] = OrderJoinString.Customize5,
    [Customize4] = OrderJoinString.Customize4,
    OrderInfo.CustomerPo,
    [ReportTypeID] = 
        case 
            when OrderInfo.IsDestJP = 1 or AQL.InspectionLevels = 2 then 'APP - AQL Outbound - Japan orders'
            when f.AcceptableQualityLevelsUkey = -1 then 'APP - AQL Outbound - 100% Inspection'
            else 'APP - AQL Outbound - Regular orders' 
        end,
    [DateStarted] = format(dateadd(hour, -system.UTCOffsetTimeZone, 
                                DATEADD(SECOND, 
								   DATEDIFF(SECOND, 0, CAST(f.AddDate AS TIME)), 
								   CAST(f.AuditDate AS DATETIME))-- 用 AuditDate 的日期覆蓋 AddDate，但保留 AddDate 的時間部分
                            ), 'yyyy-MM-ddTHH:mm:ss'),
    [InspectionCompletedDate] = format(dateadd(hour, -system.UTCOffsetTimeZone,                                 
                                DATEADD(SECOND, @RandomHours * 3600, DATEADD(SECOND, 
								   DATEDIFF(SECOND, 0, CAST(f.AddDate AS TIME)), 
								   CAST(f.AuditDate AS DATETIME))
								) --  對新的 AddDate 加上隨機 1.2 ~ 2.2 小時，並覆蓋到 EditDate
                            ), 'yyyy-MM-ddTHH:mm:ss'),
    f.OthersRemark,
    f.BAQty,
    [MeasurementResult] = cast(iif(exists(select 1 from FinalInspection_Measurement fm with (nolock) where f.ID = fm.ID), 1, 0) as bit),
    [MoistureResult] = 
        case 
            when exists (select 1 from FinalInspection_Moisture fmo with (nolock) where f.ID = fmo.ID and fmo.Result = 'F') then 'fail'
            when not exists (select 1 from FinalInspection_Moisture fmo with (nolock) where f.ID = fmo.ID) then 'na'
            else 'pass' 
        end,
    [MoistureComment] = SUBSTRING(MoistureComment.val, 1, 255),
	OrderJoinString.SewLine,
	OrderJoinString.Dest
from 
FinalInspection f with (nolock)
cross join system
outer apply (
    SELECT CustPONo = Stuff((select distinct concat(',', CustPONo) 
                        from FinalInspection_Order fo with (nolock) 
                        inner join SciProduction_Orders o with (nolock) on fo.OrderID = o.ID
                        where fo.ID = f.ID
                        FOR XML PATH('')), 1, 1, '')
           ,Customize4 = Stuff((select distinct concat(',', o.Customize4) 
                        from FinalInspection_Order fo with (nolock) 
                        inner join SciProduction_Orders o with (nolock) on fo.OrderID = o.ID
                        where fo.ID = f.ID
                        FOR XML PATH('')), 1, 1, '')
           ,Customize5 = Stuff((select distinct concat(',', o.Customize5) 
                        from FinalInspection_Order fo with (nolock) 
                        inner join SciProduction_Orders o with (nolock) on fo.OrderID = o.ID
                        where fo.ID = f.ID
                        FOR XML PATH('')), 1, 1, '')
		    ,SewLine = (
			    select TOP 1 i.Line
                from FinalInspection_Order fo with (nolock) 
                inner join SciProduction_Orders o with (nolock) on fo.OrderID = o.ID
			    inner join Inspection i with (nolock) on i.OrderID = o.ID
                where fo.ID = f.ID
		    )
		    ,Dest = Stuff((select distinct concat(',', o.Dest) 
                        from FinalInspection_Order fo with (nolock) 
                        inner join SciProduction_Orders o with (nolock) on fo.OrderID = o.ID
                        where fo.ID = f.ID
                        FOR XML PATH('')), 1, 1, '')
) OrderJoinString
outer apply (
    select	
        [POQty] = sum(o.Qty),
        [ETD_ETA] = max(o.BuyerDelivery),
        [CustomerPo] = max(o.CustCDID),
        [IsDestJP] = max(iif(o.Dest = 'JP', 1, 0))
    from Production.dbo.Orders o with (nolock)
    where o.CustPONo = f.CustPONO
) OrderInfo
outer apply (
    select 
        [val] = Stuff((select concat(';', Remark)   
                        from FinalInspection_Moisture fmo with (nolock) 
                        where f.ID = fmo.ID 
                        FOR XML PATH('')), 1, 1, '')
) MoistureComment
outer apply (
    select 
        InspectionLevels 
    from SciProduction_AcceptableQualityLevels aql 
    where aql.Ukey = f.AcceptableQualityLevelsUkey
) AQL
where f.ID = @ID 


select IsMaterialApproval      ,IsSealingSample      ,IsMetalDetection      ,IsFGWT      ,IsFGPT      ,IsTopSample      ,Is3rdPartyTestReport      ,IsPPSample
      ,IsGBTestForChina      ,IsCPSIAForYounthStytle      ,IsQRSSample      ,IsFactoryDisclaimer      ,IsA01Compliance      ,IsCPSIACompliance
      ,IsCustomerCountrySpecificCompliance 
from FinalInspectionGeneral with (nolock) where FinalInspectionID = @ID 

SELECT IsCloseShade      ,IsHandfeel      ,IsAppearance      ,IsPrintEmbDecorations      ,IsEmbellishmentPrint      ,IsEmbellishmentBonding      ,IsEmbellishmentHT
      ,IsEmbellishmentEMB      ,IsFiberContent      ,IsCareInstructions      ,IsDecorativeLabel      ,IsAdicomLabel      ,IsCountryofOrigion      ,IsSizeKey
      ,Is8FlagLabel      ,IsAdditionalLabel      ,IsIdLabel      ,IsMainLabel      ,IsSizeLabel      ,IsCareContentLabel      ,IsBrandLabel      ,IsBlueSignLabel
      ,IsLotLabel      ,IsSecurityLabel      ,IsSpecialLabel      ,IsVIDLabel      ,IsCNC      ,IsWovenlabel      ,IsTSize      ,IsCCLayout      ,IsShippingMark
      ,IsPolytagMarking      ,IsColorSizeQty      ,IsHangtag      ,IsJokerTag      ,IsWWMT      ,IsChinaCIT      ,IsPolybagSticker      ,IsUCCSticker
      ,IsPESheetMicropak      ,IsAdditionalHantage      ,IsUPCStickierHantage      ,IsGS1128Label     ,IsSecuritytag
FROM FinalInspectionCheckList with (nolock) where FinalInspectionID = @ID 

select	distinct
		oc.ColorID
from  Production.dbo.Order_ColorCombo oc with (nolock)
where oc.ID in (select POID from Production.dbo.Orders with (nolock) 
				where id in (select OrderID 
							 from FinalInspection_Order with (nolock) where ID = @ID))

WITH cte AS (
	select
        RowID = ROW_NUMBER() OVER (ORDER BY oqd.Article,oqd.SizeCode) ,  -- 想依什麼欄位排序，就放在 ORDER BY 裡
        oqd.Article,
        oqd.SizeCode,
		oqd.Qty,
		 SeqNumber= qb.LineItem
		from FinalInspection_Order fo
		inner join FinalInspection_Order_QtyShip foq on foq.ID=fo.ID
		inner join Production.dbo.Order_QtyShip_Detail oqd  ON fo.OrderID = oqd.ID  and foq.OrderID = oqd.Id and foq.Seq = oqd.Seq
		left join FinalInspection_Order_Breakdown qb on qb.FinalInspectionID = foq.Id and qb.Article = oqd.Article and qb.SizeCode = oqd.SizeCode
	where fo.ID = @ID
),
cteNulls AS (
    -- 只挑出 SeqNumber = NULL 的部分，再用 ROW_NUMBER() 幫它做連續編號
	select
        c.RowID,
        NewSeqNumber = ROW_NUMBER() OVER (ORDER BY c.RowID) * 10
	from cte c
    WHERE c.SeqNumber IS NULL
)

SELECT
    c.Article,
    c.SizeCode,
	[ShipQty] = sum(c.Qty),
    SeqNumber = COALESCE(c.SeqNumber, n.NewSeqNumber)
FROM cte c
LEFT JOIN cteNulls n
    ON c.RowID = n.RowID
group by c.Article ,c.SizeCode ,COALESCE(c.SeqNumber, n.NewSeqNumber)


select	s.StyleName,
		o.FactoryID,
		o.BrandID,
		s.CDCodeNew
into #tmpStyleInfo
from FinalInspection_Order fo with (nolock)
inner join Production.dbo.Orders o with (nolock) on o.ID = fo.OrderID
inner join Production.dbo.Style s with (nolock) on s.Ukey = o.StyleUkey
where fo.ID = @ID

declare @AdidasSAPERPCode varchar(3)

select top 1 @AdidasSAPERPCode = SUBSTRING(BrandAreaCode, 1, 3)
from SciProduction_Factory_BrandDefinition fb with (nolock)
where   exists (select 1 from #tmpStyleInfo where fb.ID = FactoryID and fb.BrandID = BrandID and (fb.CDCodeID = CDCodeNew or fb.CDCodeID = ''))
order by fb.CDCodeID desc

SELECT	[Style] =  Stuff((select concat( ';',StyleName)   from #tmpStyleInfo FOR XML PATH('')),1,1,''),
		[BrandAreaCode] = @AdidasSAPERPCode,
		[BrandAreaID] = (select Name from Production.dbo.DropDownList with (nolock) where Type = 'AdidasSAPERPCode' and ID = @AdidasSAPERPCode)

select	[DefectTypeDesc] = gdt.Description,
		[DefectCodeDesc] = gdc.Description,
        fd.GarmentDefectCodeID,
        gdc.Pivot88DefectCodeID,
		[CriticalQty] = iif(isnull(gdc.IsCriticalDefect, 0) = 1, fd.Qty, 0),
        [MajorQty] = iif(isnull(gdc.IsCriticalDefect, 0) = 0, fd.Qty, 0),
        fd.Ukey
from FinalInspection_Detail fd with (nolock)
left join Production.dbo.GarmentDefectType gdt with (nolock) on gdt.ID = fd.GarmentDefectTypeID
left join Production.dbo.GarmentDefectCode gdc with (nolock) on gdc.ID = fd.GarmentDefectCodeID
where fd.ID = @ID

select  [title] =  CONCAT(fdi.ID, '_', isnull(fd.GarmentDefectCodeID, ''), '_', fdi.Ukey),
        [full_filename] =  CONCAT(fdi.ID, '_', isnull(fd.GarmentDefectCodeID, ''), '_', fdi.Ukey, '.png'),
        [number] = ROW_NUMBER() OVER (PARTITION BY fdi.FinalInspection_DetailUkey ORDER BY fdi.Ukey),
        [comment] = isnull(fdi.Remark, ''),
        fdi.FinalInspection_DetailUkey
from SciPMSFile_FinalInspection_DetailImage fdi with (nolock)
inner join FinalInspection_Detail fd with (nolock) on fd.Ukey = fdi.FinalInspection_DetailUkey
where   fdi.ID = @ID

select  [title] =  CONCAT(fdi.ID, '_', fdi.Ukey),
        [full_filename] =  CONCAT(fdi.ID, '_', fdi.Ukey, '.png'),
        [number] = ROW_NUMBER() OVER (ORDER BY fdi.Ukey),
        [comment] = isnull(fdi.Remark, '')
from SciPMSFile_FinalInspection_OtherImage fdi with (nolock)
where   fdi.ID = @ID
order by fdi.Ukey


drop table #tmpStyleInfo
";

            return ExecuteDataSet(CommandType.Text, sqlGetData, parameter);
        }

        public List<string> GetPivot88FinalInspectionID(string finalInspectionID, bool isAutoSend = true)
        {
            SQLParameterCollection parameter = new SQLParameterCollection();
            string sqlGetData = @"
declare @FinalInspFromDateTransferToP88 date
select @FinalInspFromDateTransferToP88 = FinalInspFromDateTransferToP88 from system

select  ID
from Finalinspection with (nolock)
where   InspectionResult in ('Pass', 'Fail') and
        submitdate >= @FinalInspFromDateTransferToP88 and
        InspectionStage = 'Final' and
        exists (select 1 from Production.dbo.Orders o with (nolock) 
                where   o.CustPONo = Finalinspection.CustPONO and
                        o.BrandID in ('Adidas')
                ) and
        CustPONO <> ''
";
            if (!string.IsNullOrEmpty(finalInspectionID))
            {
                sqlGetData += " and ID = @ID";
                parameter.Add("@ID", finalInspectionID);
            }

            // 0:自動發送 1:手動發送，手動發送代表之前已經自動發送過，現在是手動重送
            sqlGetData += $" and IsExportToP88 = {(isAutoSend ? 0 : 1)}";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, parameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().Select(s => s["ID"].ToString()).ToList();
            }
            else
            {
                return new List<string>();
            }

        }

        public List<string> Get_FinalInspectionID_BrandID(string finalInspectionID)
        {
            SQLParameterCollection parameter = new SQLParameterCollection();
            string sqlGetData = @"
select DISTINCT o.BrandID
from FinalInspection_Order a
inner join  Production.dbo.Orders o on a.OrderID=o.ID
where a.id = @ID

";
            parameter.Add("@ID", DbType.String, finalInspectionID);

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, parameter);

            return dtResult.AsEnumerable().Select(s => s["BrandID"].ToString()).ToList();

        }

        public List<string> GetPivot88EndLineInspectionID(string inspectionID)
        {
            SQLParameterCollection parameter = new SQLParameterCollection();
            string sqlGetData = @"
declare @EOLInlineFromDateTransferToP88 date
select @EOLInlineFromDateTransferToP88 = EOLInlineFromDateTransferToP88 from system

select  ID
from InspectionReport with (nolock)
where   IsExportToP88 = 0 and
        InspectionReport.CustPONO <> '' and
        (AddDate >= @EOLInlineFromDateTransferToP88 or EditDate >= @EOLInlineFromDateTransferToP88) and
        exists (select 1 from Production.dbo.Orders o with (nolock) where o.CustPONo = InspectionReport.CustPONO and o.BrandID in ('Adidas'))
        and exists(select 1 from pass1 p with (nolock) where p.ID = InspectionReport.QC and p.Pivot88UserName !='')
";
            if (!string.IsNullOrEmpty(inspectionID))
            {
                sqlGetData += " and ID = @ID";
                parameter.Add("@ID", inspectionID);
            }

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, parameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().Select(s => s["ID"].ToString()).ToList();
            }
            else
            {
                return new List<string>();
            }

        }

        public string Get_FinalInspectionID_Top1_OrderID(string finalInspectionID)
        {
            SQLParameterCollection parameter = new SQLParameterCollection();
            string sqlGetData = @"
select top 1 a.OrderID
from FinalInspection_Order a 
where a.id = @ID
";
            parameter.Add("@ID", DbType.String, finalInspectionID);

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, parameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().Select(s => s["OrderID"].ToString()).FirstOrDefault();
            }
            else
            {
                return string.Empty;
            }

        }
        public List<string> GetPivot88InlineInspectionID(string inspectionID)
        {
            SQLParameterCollection parameter = new SQLParameterCollection();
            string sqlGetData = @"
declare @EOLInlineFromDateTransferToP88 date
select @EOLInlineFromDateTransferToP88 = EOLInlineFromDateTransferToP88 from system

select  ID
from InlineInspectionReport with (nolock)
where   IsExportToP88 = 0 and
        CustPONO <> '' and
        (AddDate >= @EOLInlineFromDateTransferToP88 or EditDate >= @EOLInlineFromDateTransferToP88) and
        exists (select 1 from Production.dbo.Orders o with (nolock) where o.CustPONo = InlineInspectionReport.CustPONO and o.BrandID in ('Adidas'))
        and exists(select 1 from pass1 p with (nolock) where p.ID = InlineInspectionReport.QC and p.Pivot88UserName !='')
";
            if (!string.IsNullOrEmpty(inspectionID))
            {
                sqlGetData += " and ID = @ID";
                parameter.Add("@ID", inspectionID);
            }

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, parameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().Select(s => s["ID"].ToString()).ToList();
            }
            else
            {
                return new List<string>();
            }

        }

        public void UpdateIsExportToP88(string ID, string inspectionType)
        {
            string sqlUpdateIsExportToP88 = string.Empty;

            switch (inspectionType)
            {
                case "FinalInspection":
                    sqlUpdateIsExportToP88 = $@"
    update FinalInspection set IsExportToP88 = 1 where ID = @ID
";
                    break;
                case "InlineInspection":
                    sqlUpdateIsExportToP88 = $@"
    update InlineInspectionReport set IsExportToP88 = 1, TransferTimeToPivot88 = getdate() where ID = @ID
    update SciPMSFile_InlineInspection_DetailImage set IsExportToP88 = 1 where InlineInspectionReportID = @ID
";
                    break;
                case "EndlineInspection":
                    sqlUpdateIsExportToP88 = $@"
    update InspectionReport set IsExportToP88 = 1, TransferTimeToPivot88 = getdate() where ID = @ID
    update SciPMSFile_Inspection_DetailImage set IsExportToP88 = 1 where InspectionReportID = @ID
";
                    break;
                default:
                    break;
            }

            if (string.IsNullOrEmpty(sqlUpdateIsExportToP88))
            {
                return;
            }

            SQLParameterCollection sqlPar = new SQLParameterCollection() {
                            { "@ID", DbType.String, ID }
                        };

            ExecuteNonQuery(CommandType.Text, sqlUpdateIsExportToP88, sqlPar);
        }

        public void ExecImp_EOLInlineInspectionReport()
        {
            string sqlExecImp_EOLInlineInspectionReport = "exec Imp_EOLInlineInspectionReport";
            using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.Required, TimeSpan.FromSeconds(1200)))
            {
                try
                {
                    ExecuteNonQuery(CommandType.Text, sqlExecImp_EOLInlineInspectionReport, 1200);
                    transaction.Complete();
                }
                catch (Exception ex)
                {
                    transaction.Dispose();
                    throw ex;
                }
            }
        }

        public BaseResult UpdateJunk(string ID)
        {
            SQLParameterCollection Parameter = new SQLParameterCollection()
            {
                { "@FinalInspectionID", DbType.String, ID },
            };

            string sqlCmd = "Update FinalInspection set InspectionResult = 'Junk' where ID = @FinalInspectionID and SubmitDate is null and InspectionResult = 'On-going'";
            int r = ExecuteNonQuery(CommandType.Text, sqlCmd, Parameter);
            if (r == 0)
            {
                return new BaseResult { Result = false, ErrorMessage = "Update Junk Fail" };
            }

            return new BaseResult { Result = true };
        }

        public List<FinalInspectionBasicGeneral> GetGeneralByBrand(string FinalInspectionID, string BrandID, string InspectionStage)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
                { "@FinalInspectionID", DbType.String, FinalInspectionID },
                { "@BrandID", DbType.String, BrandID },
                { "@InspectionStage", DbType.String, InspectionStage },

            };

            string cmd = $@"
select distinct fb.*
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_General fbg on fbg.BrandID = 'DEFAULT'
inner join  FinalInspectionBasicGeneral fb on fbg.BasicGeneralUkey = fb.Ukey
where fb.Junk = 0 and a.ID = @FinalInspectionID AND a.InspectionStage=@InspectionStage
";
            if (!string.IsNullOrEmpty(BrandID))
            {
                cmd = $@"

if exists(
    select * from FinalInspectionBasicBrand_General where BrandID =@BrandID
)
begin
    select DISTINCT b.*
    from FinalInspectionBasicBrand_General a
    inner join  FinalInspectionBasicGeneral b on a.BasicGeneralUkey = b.Ukey
    where a.BrandID = @BrandID AND a.InspectionStage=@InspectionStage
end
else
begin
    select distinct fb.*
    from FinalInspection a
    inner join FinalInspection_Order fo on a.ID = fo.ID
    inner join Production..Orders b on b.ID = fo.OrderID
    inner join FinalInspectionBasicBrand_General fbg on fbg.BrandID = 'DEFAULT' and fbg.InspectionStage = a.InspectionStage
    inner join  FinalInspectionBasicGeneral fb on fbg.BasicGeneralUkey = fb.Ukey
    where fb.Junk = 0 and a.ID = @FinalInspectionID
end 
";
            }

            var r = ExecuteList<FinalInspectionBasicGeneral>(CommandType.Text, cmd, objParameter);

            return r.Any() ? r.ToList() : new List<FinalInspectionBasicGeneral>();
        }
        public List<FinalInspectionBasicCheckList> GetCheckListByBrand(string FinalInspectionID, string BrandID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
                { "@FinalInspectionID", DbType.String, FinalInspectionID },
                { "@BrandID", DbType.String, BrandID },

            };

            string cmd = $@"
select distinct fb.*
from FinalInspection a
inner join FinalInspection_Order fo on a.ID = fo.ID
inner join Production..Orders b on b.ID = fo.OrderID
inner join FinalInspectionBasicBrand_CheckList fbg on fbg.BrandID = 'DEFAULT'
inner join  FinalInspectionBasicCheckList fb on fbg.BasicCheckListUkey = fb.Ukey
where fb.Junk = 0 and a.ID = @FinalInspectionID
";
            if (!string.IsNullOrEmpty(BrandID))
            {
                cmd = $@"
if exists(
    select * from FinalInspectionBasicBrand_CheckList where BrandID =@BrandID
)
begin
    select DISTINCT b.*
    from FinalInspectionBasicBrand_CheckList a
    inner join  FinalInspectionBasicCheckList b on a.BasicCheckListUkey = b.Ukey
    where a.BrandID = @BrandID

end
else
begin
    select distinct fb.*
    from FinalInspection a
    inner join FinalInspection_Order fo on a.ID = fo.ID
    inner join Production..Orders b on b.ID = fo.OrderID
    inner join FinalInspectionBasicBrand_CheckList fbg on fbg.BrandID = 'DEFAULT'
    inner join  FinalInspectionBasicCheckList fb on fbg.BasicCheckListUkey = fb.Ukey
    where fb.Junk = 0 and a.ID = @FinalInspectionID
end
";
            }

            var r = ExecuteList<FinalInspectionBasicCheckList>(CommandType.Text, cmd, objParameter);

            return r.Any() ? r.ToList() : new List<FinalInspectionBasicCheckList>();
        }


        public List<FinalInspectionBasicGeneral> GetAllGeneral()
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string cmd = $@"
select *
from FinalInspectionBasicGeneral
";
            var r = ExecuteList<FinalInspectionBasicGeneral>(CommandType.Text, cmd, objParameter);

            return r.Any() ? r.ToList() : new List<FinalInspectionBasicGeneral>();
        }
        public List<FinalInspectionBasicCheckList> GetAllCheckList()
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string cmd = $@"
select *
from FinalInspectionBasicCheckList
";

            var r = ExecuteList<FinalInspectionBasicCheckList>(CommandType.Text, cmd, objParameter);

            return r.Any() ? r.ToList() : new List<FinalInspectionBasicCheckList>();
        }


        public List<FinalInspectionSignature> GetFinalInspectionSignature(FinalInspectionSignature Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@finalInspectionID", DbType.String, Req.FinalInspectionID },
            { "@JobTitle", DbType.String, Req.JobTitle },
            { "@UserID", DbType.String, Req.UserID }
            };

            string cmd = @"
select a.FinalInspectionID
		,a.UserID
		,UserName = p.Name
		,a.JobTitle
		,b.Signature
		,b.AddName
		,b.AddDate
from  FinalInspectionSignature a
INNER JOIN Pass1 p on a.UserID=p.ID
left join PMSFile.dbo. FinalInspectionSignature b on a.FinalInspectionID = b.FinalInspectionID AND a.JobTitle = b.JobTitle AND a.UserID = b.UserID
where  a.FinalInspectionID = @finalInspectionID
";

            if (!string.IsNullOrEmpty(Req.JobTitle))
            {
                cmd += " and a.JobTitle = @JobTitle";
            }
            if (!string.IsNullOrEmpty(Req.UserID))
            {
                cmd += " and a.UserID = @UserID";
            }
            var r = ExecuteList<FinalInspectionSignature>(CommandType.Text, cmd, objParameter);

            return r.Any() ? r.ToList() : new List<FinalInspectionSignature>();
        }
        public bool InsertFinalInspectionSignature(FinalInspectionSignature Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@finalInspectionID", DbType.String, Req.FinalInspectionID },
            { "@JobTitle", DbType.String, Req.JobTitle },
            { "@UserID", DbType.String, Req.UserID },
            { "@AddName", DbType.String, Req.AddName },
            { "@Signature", Req.Signature },
            };

            string cmd = @"
if not exists(
    select 1 
    from PMSFile.dbo.FinalInspectionSignature
    where FinalInspectionID = @finalInspectionID AND JobTitle = @JobTitle AND UserID = @UserID
)
BEGIN
    INSERT INTO PMSFile.dbo.FinalInspectionSignature
               (FinalInspectionID,UserID,JobTitle,Signature,AddName,AddDate)
    VALUES
               (@FinalInspectionID ,@UserID ,@JobTitle ,@Signature ,@AddName ,GETDATE() )
END
ELSE
BEGIN
    UPDATE PMSFile.dbo.FinalInspectionSignature
    SET Signature = @Signature ,AddName = @AddName  ,AddDate = GETDATE()  
    where FinalInspectionID = @finalInspectionID AND JobTitle = @JobTitle AND UserID = @UserID
END


";

            ExecuteNonQuery(CommandType.Text, cmd, objParameter);
            return true;
        }

        public List<FinalInspectionSignature> GetFinalInspectionSignatureUser()
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string cmd = @"
select UserID= a.ID , UserName=b.Name
from Quality_Pass1 a
inner join Pass1 b on a.ID=b.ID
where a.Junk=0
";
            var r = ExecuteList<FinalInspectionSignature>(CommandType.Text, cmd, objParameter);

            return r.Any() ? r.ToList() : new List<FinalInspectionSignature>();
        }

        public bool InsertFinalInspectionSignatureUser(string FinalInspectionID, string JobTitle, List<FinalInspectionSignature> allData)
        {
            List<FinalInspectionSignature> oldData = this.GetFinalInspectionSignature(new FinalInspectionSignature()
            {
                FinalInspectionID = FinalInspectionID,
                JobTitle = JobTitle
            }).ToList();

            List<FinalInspectionSignature> needUpdateDetailList =
                PublicClass.CompareListValue<FinalInspectionSignature>(
                    allData,
                    oldData,
                    "FinalInspectionID,UserID,JobTitle",
                    "FinalInspectionID,UserID,JobTitle");

            string insertDetail = $@" ----寫入 
INSERT INTO FinalInspectionSignature (FinalInspectionID,UserID,JobTitle)
VALUES  (@FinalInspectionID ,@UserID ,@JobTitle)
";

            string deleteDetail = $@" ----刪除 
DELETE FROM FinalInspectionSignature where FinalInspectionID = @FinalInspectionID AND UserID = @UserID AND JobTitle = @JobTitle
DELETE FROM PMSFile.dbo.FinalInspectionSignature where FinalInspectionID = @FinalInspectionID AND UserID = @UserID AND JobTitle = @JobTitle
";

            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();

                listDetailPar.Add(new SqlParameter($"@FinalInspectionID", detailItem.FinalInspectionID));
                listDetailPar.Add(new SqlParameter($"@UserID", detailItem.UserID));
                listDetailPar.Add(new SqlParameter($"@JobTitle", detailItem.JobTitle));
                switch (detailItem.StateType)
                {
                    case CompareStateType.Add:
                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);
                        break;
                    case CompareStateType.Delete:
                        ExecuteNonQuery(CommandType.Text, deleteDetail, listDetailPar);
                        break;
                    case CompareStateType.None:
                        break;
                    default:
                        break;
                }
            }
            return true;
        }
    }
}
