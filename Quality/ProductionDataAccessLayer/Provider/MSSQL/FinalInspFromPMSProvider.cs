using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Data.SqlClient;
using DatabaseObject.ViewModel.FinalInspection;
using System.Linq;
using ToolKit;
using System.Web.Mvc;
using DatabaseObject.ManufacturingExecutionDB;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class FinalInspFromPMSProvider : SQLDAL, IFinalInspFromPMSProvider
    {
        #region 底層連線
        public FinalInspFromPMSProvider(string ConString) : base(ConString) { }
        public FinalInspFromPMSProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<SelectSewing> GetSelectedSewingLine(string FactoryID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            //台北
            SbSql.Append($@"
Select SewingLine = ID
From Production.dbo.SewingLine --工廠
Where Junk = 0
AND FactoryID=@FactoryID
");

            if (!string.IsNullOrEmpty(FactoryID))
            {
                SbSql.Append($@"AND FactoryID  = @FactoryID ");

                paras.Add("@FactoryID", DbType.String, FactoryID);
            }


            return ExecuteList<SelectSewing>(CommandType.Text, SbSql.ToString(), paras);
        }
        public IList<SelectSewing> GetSelectedSewingLineFromEndline(List<string> listOrderID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            string whereOrderID = listOrderID.Select(s => $"'{s}'").JoinToString(",");
            //台北
            SbSql.Append($@"
select distinct SewingLine = Line
from Inspection
where OrderID  in ({whereOrderID})
");

            return ExecuteList<SelectSewing>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<SelectSewingTeam> GetSelectedSewingTeam()
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            //台北
            SbSql.Append($@"
Select SewingTeamID = ID
from Production.dbo.SewingTeam  WITH(NOLOCK)
Where Junk = 0
");


            return ExecuteList<SelectSewingTeam>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<SelectedPO> GetSelectedPOForInspection(List<string> listOrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string whereOrderID = listOrderID.Select(s => $"'{s}'").JoinToString(",");
            string sqlGetData = $@"
select  [OrderID] = o.id,
        o.CustPONO,
        o.StyleID,
        o.SeasonID,
        o.BrandID,
        o.Qty,
        [AvailableQty] = 0,
        [Cartons] = '',
        [Seq] = '',
        [Article] = (SELECT Stuff((select concat( ',',Article)   from Order_Article WITH(NOLOCK) where ID = o.ID FOR XML PATH('')),1,1,'') ),
		MetalContaminateQty = ISNULL(firstMD.ErrQty,0) + ISNULL(secondMD.ErrQty,0)
  from  Orders o with (nolock)
outer apply(
	 select ErrQty = SUM(ErrQty)
	 from Production.dbo.PackErrTransfer pe 
	 where pe.OrderID = o.ID 
)firstMD
outer apply(
	 select ErrQty = SUM(ErrQty)
	 from Production.dbo.ClogPackingError  cp 
	 where cp.OrderID = o.ID 
)secondMD
 where  o.id in ({whereOrderID})
";
            return ExecuteList<SelectedPO>(CommandType.Text, sqlGetData, listPar);
        }

        public IList<SelectedPO> GetSelectedPOForInspection(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetData = $@"
select  OrderID,
        AvailableQty
into    #FinalInspection_Order
from    ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where   ID  =   @finalInspectionID

select  [OrderID] = o.id,
        o.CustPONO,
        o.StyleID,
        o.SeasonID,
        o.BrandID,
        [Qty] = 0,
        [AvailableQty] = fo.AvailableQty,
		MetalContaminateQty = ISNULL(firstMD.ErrQty,0) + ISNULL(secondMD.ErrQty,0),
        [Cartons] = ''
from  Production.dbo.Orders o with (nolock)
inner join  #FinalInspection_Order fo on fo.OrderID = o.ID
outer apply(
	 select ErrQty = SUM(ErrQty)
	 from MainServer.Production.dbo.PackErrTransfer pe 
	 where pe.OrderID = fo.OrderID 
)firstMD
outer apply(
	 select ErrQty = SUM(ErrQty)
	 from MainServer.Production.dbo.ClogPackingError cp 
	 where cp.OrderID = fo.OrderID 
)secondMD

drop table #FinalInspection_Order
";
            return ExecuteList<SelectedPO>(CommandType.Text, sqlGetData, listPar);
        }

        public IList<SelectCarton> GetSelectedCartonForSetting(List<string> listOrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string whereOrderID = listOrderID.Select(s => $"'{s}'").JoinToString(",");
            string sqlGetData = $@"
select  [Selected] = Cast(0 as bit),
        pld.OrderID,
        [PackingListID] = pld.id, 
        [CTNNo] = CTNStartNo,
        [Seq] = pld.OrderShipmodeSeq
		,ShipQty = SUM(pld.ShipQty)
 from PackingList_Detail pld WITH(NOLOCK)
 where  pld.OrderID in ({whereOrderID})  
    and CTNStartNo <> ''
    --and CTNQty = 1
 group by  OrderID,ID,CTNStartNo,OrderShipmodeSeq
 ORDER BY CTNStartNo
";
            return ExecuteList<SelectCarton>(CommandType.Text, sqlGetData, listPar);
        }

        public IList<SelectCarton> GetSelectedCartonForSetting(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetData = $@"
select  [Selected] = 1,
        OrderID,
        PackinglistID,
        CTNNo,
        Seq
into    #FinalInspection_OrderCarton
from    ManufacturingExecution.dbo.FinalInspection_OrderCarton with (nolock)
where   ID  =   @finalInspectionID

select  OrderID
into    #FinalInspection_Order
from    ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where   ID  =   @finalInspectionID

select  [Selected] = cast(isnull(fc.Selected, 0) as bit),
        pld.OrderID,
        [PackingListID] = pld.id, 
        [CTNNo] = CTNStartNo,
        [Seq] = pld.OrderShipmodeSeq
from MainServer.Production.dbo.PackingList_Detail pld WITH(NOLOCK)
left join   #FinalInspection_OrderCarton fc on  fc.OrderID = pld.OrderID and 
                                                fc.PackinglistID = pld.ID and 
                                                fc.CTNNo = pld.CTNStartNo and
                                                fc.Seq = pld.OrderShipmodeSeq
where   pld.OrderID in (select OrderID from #FinalInspection_Order) and
        pld.CTNQty = 1

";
            return ExecuteList<SelectCarton>(CommandType.Text, sqlGetData, listPar);
        }

        public IList<SelectOrderShipSeq> GetSelectOrderShipSeqForSetting(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetData = $@"
select  [Selected] = 1,
        OrderID,
        Seq,
        ShipmodeID
into    #FinalInspection_Order_QtyShip
from    ManufacturingExecution.dbo.FinalInspection_Order_QtyShip with (nolock)
where   ID  =   @finalInspectionID

select  OrderID
into    #FinalInspection_Order
from    ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where   ID  =   @finalInspectionID

select  [Selected] = cast(isnull(foq.Selected, 0) as bit),
        [OrderID] = oqs.ID,
        [Seq] = oqs.Seq, 
        [ShipmodeID] = oqs.ShipmodeID,
        [Article] = (SELECT Stuff((select distinct concat( ',',Article)   
                                    from Production.dbo.Order_QtyShip_Detail with (nolock) 
                                    where ID = oqs.ID and Seq = oqs.Seq FOR XML PATH('')),1,1,'') ),
        [Qty] = oqs.Qty
from Production.dbo.Order_QtyShip oqs with (nolock)
left join   #FinalInspection_Order_QtyShip foq on   foq.OrderID = oqs.ID and 
                                                    foq.Seq = oqs.Seq 
where   oqs.ID in (select OrderID from #FinalInspection_Order)

";
            return ExecuteList<SelectOrderShipSeq>(CommandType.Text, sqlGetData, listPar);
        }

        public IList<AcceptableQualityLevels> GetAcceptableQualityLevelsForSetting()
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlGetData = $@"


select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
INTO #AllData
from AcceptableQualityLevels WITH(NOLOCK)
where  Junk = 0 and  AQLType in (1,1.5,2.5) and InspectionLevels IN ('1','2') and AcceptedQty is not null 
AND BrandID='' AND Category = ''
UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (1.5,2.5,4.0) and InspectionLevels IN ('1') and AcceptedQty is not null 
AND BrandID='U.ARMOUR' AND Category = ''

UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (1.0,1.5,2.5) and InspectionLevels IN ('1')  and AcceptedQty is not null 
AND BrandID='NIKE' AND Category = ''

UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (1.5) and InspectionLevels IN ('1','2')  and AcceptedQty is not null 
AND BrandID='LLL' AND Category = ''

UNION

select	BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (2.5) and InspectionLevels IN ('2')  and AcceptedQty is not null 
AND BrandID='N.FACE' AND Category = ''

UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (2.5) and InspectionLevels IN ('2')  and AcceptedQty is not null 
AND BrandID='GYMSHARK' AND Category = ''

UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where Junk = 0 and AQLType in (1) and InspectionLevels IN ('S-4') 
and AcceptedQty is not null 
AND BrandID='REI' AND Category = ''
UNION 

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where Junk = 0 and AQLType in (1.5) and InspectionLevels IN ('1') 
and AcceptedQty is not null 
AND BrandID='Kolon' AND Category = ''
UNION 

select BrandID='AllBrand',AQLType=100,InspectionLevels='100% Inspection'
,LotSize_Start=0,LotSize_End=0,SampleSize=0,AcceptedQty=0,Ukey=0

order by BrandID,AQLType , InspectionLevels

select *
from #AllData

drop table #AllData
";
            return ExecuteList<AcceptableQualityLevels>(CommandType.Text, sqlGetData, listPar);
        }

        public IList<AcceptableQualityLevelsProList> GetAcceptableQualityLevelsProListForSetting(string BrandID ,long Ukey)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@BrandID", BrandID);

            string sqlGetData = $@"
select a.ProUkey
    ,a.BrandID
    ,a.InspectionLevels
    ,a.AQLType
    ,a.LotSize_Start
    ,a.LotSize_End
    ,a.SampleSize
    ,b.AQLDefectCategoryUkey
    ,c.Description,b.AcceptedQty
    ,DefectDescription = c.Description
from SciProduction_AcceptableQualityLevelsPro a WITH(NOLOCK)
inner join SciProduction_AcceptableQualityLevelsPro_Detail b WITH(NOLOCK) on a.ProUkey=b.ProUkey
inner join SciProduction_AcceptableQualityLevelsPro_DefectCategory c WITH(NOLOCK) on b.AQLDefectCategoryUkey=c.Ukey
where  Junk = 0
AND BrandID = @BrandID AND Category = ''
";
            if (Ukey > 0)
            {
                listPar.Add("@Ukey", Ukey);
                sqlGetData += "AND a.ProUkey = @Ukey";
            }
            return ExecuteList<AcceptableQualityLevelsProList>(CommandType.Text, sqlGetData, listPar);
        }
        public IList<AcceptableQualityLevels> GetAcceptableQualityLevelsForMeasurement()
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlGetData = $@"


select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
INTO #AllData
from AcceptableQualityLevels WITH(NOLOCK)
where  Junk = 0 and  AQLType in (1,1.5,2.5) and InspectionLevels IN ('1','2') and AcceptedQty is not null 
AND BrandID='' AND Category = 'Measurement'
UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (1.5,2.5,4.0) and InspectionLevels IN ('1') and AcceptedQty is not null 
AND BrandID='U.ARMOUR' AND Category = 'Measurement'

UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (1.0,1.5,2.5) and InspectionLevels IN ('1')  and AcceptedQty is not null 
AND BrandID='NIKE' AND Category = 'Measurement'

UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (1.5) and InspectionLevels IN ('1')  and AcceptedQty is not null 
AND BrandID='LLL' AND Category = 'Measurement'

UNION

select	BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (2.5) and InspectionLevels IN ('2')  and AcceptedQty is not null 
AND BrandID='N.FACE' AND Category = 'Measurement'

UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where  Junk = 0 and  AQLType in (2.5) and InspectionLevels IN ('2')  and AcceptedQty is not null 
AND BrandID='GYMSHARK' AND Category = 'Measurement'

UNION

select BrandID,AQLType,InspectionLevels ,LotSize_Start,LotSize_End,SampleSize,AcceptedQty,Ukey
from AcceptableQualityLevels
where Junk = 0 and AQLType in (1) and InspectionLevels IN ('S-4') 
and AcceptedQty is not null 
AND BrandID='REI' AND Category = 'Measurement'
UNION 
select BrandID='AllBrand',AQLType=100,InspectionLevels='100% Inspection'
,LotSize_Start=0,LotSize_End=0,SampleSize=0,AcceptedQty=0,Ukey=0

order by BrandID,AQLType , InspectionLevels

select *
from #AllData

drop table #AllData
";
            return ExecuteList<AcceptableQualityLevels>(CommandType.Text, sqlGetData, listPar);
        }
        public IList<FinalInspectionDefectItem> GetFinalInspectionDefectItems(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetData = $@"
select  GarmentDefectTypeID,
        GarmentDefectCodeID,
        Qty,
        Ukey,
        AreaCode,
        Remark,
		Operation = Operation.Operation,
		Operator = Operation.Operator,
		OperatorText = Operation.OperatorText
into #FinalInspection_Detail
from ManufacturingExecution.dbo.FinalInspection_Detail a
outer apply(
	select Operation= STUFF((
			select ',' + b.InlineOperation
			from ManufacturingExecution.dbo.FinalInspection_Detail_Operation b
			where b.InspectionDetailUkey=a.Ukey
			for xml path('')
		),1,1,'')
		,Operator= STUFF((
			select ',' + b.InlineOperator
			from ManufacturingExecution.dbo.FinalInspection_Detail_Operation b
			where b.InspectionDetailUkey=a.Ukey
			for xml path('')
		),1,1,'')
		,OperatorText= STUFF((
			select ',' + c.FirstName+' '+c.LastName
			from FinalInspection_Detail_Operation b
			inner join InlineEmployee c on c.EmployeeID = b.InlineOperator
			where b.InspectionDetailUkey=a.Ukey
			for xml path('')
		),1,1,'')
)Operation
where   ID = @finalInspectionID

select  [Ukey] = isnull(fd.Ukey, -1),
        [DefectType] = gdt.ID,
        [DefectCode] = gdc.ID,
        [DefectTypeDesc] = gdt.ID +'-'+gdt.Description,
        [DefectCodeDesc] = gdc.ID +'-'+gdc.Description,
        [Qty] = isnull(fd.Qty, 0),
        Operation = ISNULL( fd.Operation ,''),
        Operator = ISNULL( fd.Operator ,''),
        OperatorText = ISNULL( fd.OperatorText ,''),
        Remark = ISNULL( fd.Remark ,''),
		[RowIndex]=ROW_NUMBER() OVER(ORDER BY gdt.id,gdc.id) -1
		,HasImage = Cast(
			IIF(EXISTS(
				select 1 from SciPMSFile_FinalInspection_DetailImage img 
				where img.FinalInspection_DetailUkey = isnull(fd.Ukey, -1)
			),1,0)		
		as bit),
        AreaCode = ISNULL( fd.AreaCode ,'')
    from [MainServer].Production.dbo.GarmentDefectType gdt with (nolock)
    inner join [MainServer].Production.dbo.GarmentDefectCode gdc with (nolock) on gdt.id=gdc.GarmentDefectTypeID
    left join   #FinalInspection_Detail fd on fd.GarmentDefectTypeID = gdt.ID and fd.GarmentDefectCodeID = gdc.ID
    where   gdt.Junk =0 and
            gdc.Junk =0
 order by gdt.id,gdc.id

 drop table #FinalInspection_Detail
";
            return ExecuteList<FinalInspectionDefectItem>(CommandType.Text, sqlGetData, listPar);
        }

        public IList<FinalInspection_DefectDetail> GetFinalInspection_DefectDetails(string finalInspectionID,long ProUkey)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@FinalInspectionID", finalInspectionID);
            listPar.Add("@ProUkey", ProUkey);

            string sqlGetData = $@"
select a.ProUkey
	,a.BrandID
	,DefectCategoryDescription = c.Description
    ,b.AcceptedQty  
	,DefectCategoryUkey = c.Ukey
	,defect.DefectQty
	,DefectCategoryResult = IIF(b.AcceptedQty < defect.DefectQty,'Pass','Fail')
from SciProduction_AcceptableQualityLevelsPro a WITH(NOLOCK)
inner join SciProduction_AcceptableQualityLevelsPro_Detail b on a.ProUkey=b.ProUkey
inner join SciProduction_AcceptableQualityLevelsPro_DefectCategory c on b.AQLDefectCategoryUkey=c.Ukey
OUTER APPLY(
	select top 1 DefectQty
	from ManufacturingExecution..FinalInspection_DefectDetail d
	where  d.DefectCategoryUkey=c.Ukey and d.ProUkey=a.ProUkey and d.FinalInspectionID = @FinalInspectionID
)defect
where a.ProUkey = @ProUkey
";
            return ExecuteList<FinalInspection_DefectDetail>(CommandType.Text, sqlGetData, listPar);
        }
        public List<string> GetArticleList(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetMoistureArticleList = @"
select  OrderID, Seq
into #FinalInspection_Order_QtyShip
from ManufacturingExecution.dbo.FinalInspection_Order_QtyShip with (nolock)
where ID = @finalInspectionID

select distinct oqd.Article 
from Production.dbo.Order_QtyShip_Detail oqd with (nolock)
where exists (select 1 from #FinalInspection_Order_QtyShip where OrderID = oqd.ID and Seq = oqd.Seq )
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetMoistureArticleList, listPar);

            if (dtResult.Rows.Count == 0)
            {
                return new List<string>();
            }
            else
            {
                return dtResult.AsEnumerable().Select(s => s["Article"].ToString()).ToList();
            }

        }

        public IList<SelectListItem> GetActionSelectListItem()
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            string sqlGetActionSelectListItem = @"
select  [Text] = '', [Value] = ''
union
select  [Text] = Name, [Value] = Name 
from Production.dbo.DropDownList ddl WITH(NOLOCK) where
type='PMS_MoistureAction'

";
            return ExecuteList<SelectListItem>(CommandType.Text, sqlGetActionSelectListItem, listPar);
        }

        public IList<ArticleSize> GetArticleSizeList(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetMoistureArticleList = @"
select  OrderID, Seq
into #FinalInspection_Order_QtyShip
from ManufacturingExecution.dbo.FinalInspection_Order_QtyShip with (nolock)
where ID = @finalInspectionID

select distinct oqd.Article, oqd.SizeCode 
from Production.dbo.Order_QtyShip_Detail oqd with (nolock)
where exists (select 1 from #FinalInspection_Order_QtyShip where OrderID = oqd.ID and Seq = oqd.Seq )
";

            return ExecuteList<ArticleSize>(CommandType.Text, sqlGetMoistureArticleList, listPar);
        }

        public List<string> GetProductTypeList(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetMoistureArticleList = @"
select OrderID
into #FinalInspection_Order
from ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where ID = @finalInspectionID

----避免沒有Order_Location資料，預先塞入
DECLARE CUR_SewingOutput_Detail CURSOR FOR 
     Select distinct orderid from #FinalInspection_Order

declare @orderid varchar(13) 
OPEN CUR_SewingOutput_Detail   
FETCH NEXT FROM CUR_SewingOutput_Detail INTO @orderid 
WHILE @@FETCH_STATUS = 0 
BEGIN
  exec MainServer.Production.dbo.Ins_OrderLocation @orderid
FETCH NEXT FROM CUR_SewingOutput_Detail INTO @orderid
END
CLOSE CUR_SewingOutput_Detail
DEALLOCATE CUR_SewingOutput_Detail


select distinct Location 
from Production.dbo.Order_Location with (nolock)
where OrderId in (select OrderID from #FinalInspection_Order)
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetMoistureArticleList, listPar);

            if (dtResult.Rows.Count == 0)
            {
                return new List<string>();
            }
            else
            {
                return dtResult.AsEnumerable().Select(s => s["Location"].ToString()).ToList();
            }
        }
        public int GetMeasurementRemainingAmount(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetMoistureArticleList = @"
DECLARE @SampleSize as int =(
	select IIF(MeasurementSampleSize = 0, SampleSize ,MeasurementSampleSize)
	from    FinalInspection WITH(NOLOCK)
	where   ID = @finalInspectionID
)

DECLARE @MeasurementCount as int=(
	SELECT COUNT(1) FROM(
		select DISTINCT Article,SizeCode,Location,AddDate
		from    FinalInspection_Measurement WITH(NOLOCK)
		where   ID = @finalInspectionID
	) a
)
SELECT RemainingAmount = @SampleSize - @MeasurementCount
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetMoistureArticleList, listPar);

            if (dtResult.Rows.Count == 0)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt16(dtResult.Rows[0]["RemainingAmount"]) < 0 ? 0 : Convert.ToInt16(dtResult.Rows[0]["RemainingAmount"]);
            }
        }

        public IList<SelectOrderShipSeq> GetSelectOrderShipSeqForSetting(List<string> listOrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string whereOrderID = listOrderID.Select(s => $"'{s}'").JoinToString(",");
            string sqlGetData = $@"
select  [Selected] = cast(0 as bit),
        [OrderID] = oqs.ID,
        [Seq] = oqs.Seq, 
        [ShipmodeID] = oqs.ShipmodeID,
        [Article] = (SELECT Stuff((select distinct concat( ',',Article)   
                                    from Order_QtyShip_Detail with (nolock) 
                                    where ID = oqs.ID and Seq = oqs.Seq FOR XML PATH('')),1,1,'') ),
        [Qty] = oqs.Qty
from Order_QtyShip oqs with (nolock)
where   oqs.id in ({whereOrderID})
";
            return ExecuteList<SelectOrderShipSeq>(CommandType.Text, sqlGetData, listPar);
        }

        public IList<DatabaseObject.ProductionDB.System> GetSystem()
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

        public void UpdateOrderQtyShip(string finalInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@FinalInspectionID", finalInspectionID }
            };

            string sqlUpdCmd = @"
update q
	set q.CFAUpdateDate = t.SubmitDate
	, q.[CFAFinalInspectResult] = iif(t2.InspectionStage = 'Final', t2.InspectionResult, q.CFAFinalInspectResult )
	, q.[CFAFinalInspectDate] = iif(t2.InspectionStage = 'Final', t2.AuditDate, q.CFAFinalInspectDate)
	, q.[CFAFinalInspectHandle] = iif(t2.InspectionStage = 'Final', t2.CFA, q.CFAFinalInspectHandle )
	, q.[CFA3rdInspectResult] = iif(t2.InspectionStage = '3rd Party', t2.InspectionResult, q.CFA3rdInspectResult )
	, q.[CFA3rdInspectDate] = iif(t2.InspectionStage = '3rd Party', t2.AuditDate, q.CFA3rdInspectDate)
	, q.[CFAIs3rdInspectHandle] = iif(t2.InspectionStage = '3rd Party', t2.CFA, q.CFAIs3rdInspectHandle )
from Order_QtyShip q
inner join (
	select foq.OrderID, foq.Seq
		, f.SubmitDate
	from [ExtendServer].ManufacturingExecution.dbo.FinalInspection f
	inner join [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order_QtyShip foq on f.ID = foq.ID
	inner join (
		select [InspectionTimes] = MAX(foq.InspectionTimes), foq.OrderID, foq.Seq
		from [ExtendServer].ManufacturingExecution.dbo.FinalInspection f
		inner join [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order_QtyShip foq on f.ID = foq.ID
		where exists (
			select 1 
			from [ExtendServer].ManufacturingExecution.dbo.FinalInspection f2
			inner join [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order_QtyShip foq2 on f2.ID = foq2.ID
			where f2.ID = @FinalInspectionID
			and foq.OrderID = foq2.OrderID 
			and foq.Seq = foq2.Seq)
		and f.InspectionStep = 'Submit' 
		and f.SubmitDate is not null
		group by foq.OrderID, foq.Seq
	) t on foq.OrderID = t.OrderID and foq.Seq = t.Seq and foq.InspectionTimes = t.InspectionTimes
)t on q.Id = t.OrderID and q.Seq = t.Seq
left join (
	select foq.OrderID, foq.Seq
		, f.InspectionStage
		, f.InspectionResult
		, f.AuditDate
		, f.CFA
	from [ExtendServer].ManufacturingExecution.dbo.FinalInspection f
	inner join [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order_QtyShip foq on f.ID = foq.ID
	inner join (
		select [InspectionTimes] = MAX(foq.InspectionTimes), f.InspectionStage, foq.OrderID, foq.Seq
		from [ExtendServer].ManufacturingExecution.dbo.FinalInspection f
		inner join [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order_QtyShip foq on f.ID = foq.ID
		where exists (
			select 1 
			from [ExtendServer].ManufacturingExecution.dbo.FinalInspection f2
			inner join [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order_QtyShip foq2 on f2.ID = foq2.ID
			where f2.ID = @FinalInspectionID
			and foq.OrderID = foq2.OrderID 
			and foq.Seq = foq2.Seq)
		and f.InspectionStep = 'Submit' 
		and f.SubmitDate is not null
		and f.InspectionStage in ('Final', '3rd Party')
		group by f.InspectionStage, foq.OrderID, foq.Seq
	) t on foq.OrderID = t.OrderID and foq.Seq = t.Seq and foq.InspectionTimes = t.InspectionTimes and f.InspectionStage = t.InspectionStage
)t2 on q.Id = t2.OrderID and q.Seq = t2.Seq

";
            ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }
    }
}
