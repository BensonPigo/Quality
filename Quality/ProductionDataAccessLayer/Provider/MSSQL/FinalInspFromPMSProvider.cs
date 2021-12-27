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
        [Article] = (SELECT Stuff((select concat( ',',Article)   from Order_Article WITH(NOLOCK) where ID = o.ID FOR XML PATH('')),1,1,'') )
  from  Orders o with (nolock)
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
from    [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where   ID  =   @finalInspectionID

select  [OrderID] = o.id,
        o.CustPONO,
        o.StyleID,
        o.SeasonID,
        o.BrandID,
        [Qty] = 0,
        [AvailableQty] = fo.AvailableQty,
        [Cartons] = ''
from  Orders o with (nolock)
inner join  #FinalInspection_Order fo on fo.OrderID = o.ID
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
    --and CTNQty = 1
 group by  OrderID,ID,CTNStartNo,OrderShipmodeSeq
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
from    [ExtendServer].ManufacturingExecution.dbo.FinalInspection_OrderCarton with (nolock)
where   ID  =   @finalInspectionID

select  OrderID
into    #FinalInspection_Order
from    [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where   ID  =   @finalInspectionID

select  [Selected] = cast(isnull(fc.Selected, 0) as bit),
        pld.OrderID,
        [PackingListID] = pld.id, 
        [CTNNo] = CTNStartNo,
        [Seq] = pld.OrderShipmodeSeq
from PackingList_Detail pld WITH(NOLOCK)
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
from    [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order_QtyShip with (nolock)
where   ID  =   @finalInspectionID

select  OrderID
into    #FinalInspection_Order
from    [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where   ID  =   @finalInspectionID

select  [Selected] = cast(isnull(foq.Selected, 0) as bit),
        [OrderID] = oqs.ID,
        [Seq] = oqs.Seq, 
        [ShipmodeID] = oqs.ShipmodeID,
        [Article] = (SELECT Stuff((select distinct concat( ',',Article)   
                                    from Order_QtyShip_Detail with (nolock) 
                                    where ID = oqs.ID and Seq = oqs.Seq FOR XML PATH('')),1,1,'') ),
        [Qty] = oqs.Qty
from Order_QtyShip oqs with (nolock)
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
select	InspectionLevels ,
		LotSize_Start	 ,
		LotSize_End		 ,
		SampleSize		 ,
		Ukey			 ,
		Junk			 ,
		AQLType			 ,
		AcceptedQty
from AcceptableQualityLevels WITH(NOLOCK)
where AQLType in (1,1.5,2.5) and InspectionLevels < 3 and AcceptedQty is not null 
order by AQLType , InspectionLevels
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
        Ukey
into #FinalInspection_Detail
from [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Detail
where   ID = @finalInspectionID

select  [Ukey] = isnull(fd.Ukey, -1),
        [DefectType] = gdt.ID,
        [DefectCode] = gdc.ID,
        [DefectTypeDesc] = gdt.ID +'-'+gdt.Description,
        [DefectCodeDesc] = gdc.ID +'-'+gdc.Description,
        [Qty] = isnull(fd.Qty, 0),
		[RowIndex]=ROW_NUMBER() OVER(ORDER BY gdt.id,gdc.id) -1,
		HasImage = Cast(IIF(img.Image is null,0,1) as bit)
    from GarmentDefectType gdt with (nolock)
    inner join GarmentDefectCode gdc with (nolock) on gdt.id=gdc.GarmentDefectTypeID
    left join   #FinalInspection_Detail fd on fd.GarmentDefectTypeID = gdt.ID and fd.GarmentDefectCodeID = gdc.ID
    left join [ExtendServer].PMSFile.dbo.FinalInspection_DetailImage img on img.FinalInspection_DetailUkey = isnull(fd.Ukey, -1)
    where   gdt.Junk =0 and
            gdc.Junk =0
 order by gdt.id,gdc.id


";
            return ExecuteList<FinalInspectionDefectItem>(CommandType.Text, sqlGetData, listPar);
        }

        public List<string> GetArticleList(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetMoistureArticleList = @"
select  OrderID, Seq
into #FinalInspection_Order_QtyShip
from [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order_QtyShip with (nolock)
where ID = @finalInspectionID

select distinct oqd.Article 
from Order_QtyShip_Detail oqd with (nolock)
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
from DropDownList ddl WITH(NOLOCK) where
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
from [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order_QtyShip with (nolock)
where ID = @finalInspectionID

select distinct oqd.Article, oqd.SizeCode 
from Order_QtyShip_Detail oqd with (nolock)
where exists (select 1 from #FinalInspection_Order_QtyShip where OrderID = oqd.ID and Seq = oqd.Seq )
";

            return ExecuteList<ArticleSize>(CommandType.Text, sqlGetMoistureArticleList, listPar);
        }

        public List<string> GetProductTypeList(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetMoistureArticleList = @"
select  OrderID
into #FinalInspection_Order
from [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where ID = @finalInspectionID

----避免沒有Order_Location資料，預先塞入
INSERT into  Order_Location(OrderId,Location,Rate,AddName,AddDate,EditName,EditDate)
SELECT o.id,sl.Location,sl.Rate,sl.AddName,sl.AddDate,sl.EditName,sl.EditDate
FROM orders o WITH(NOLOCK)
inner join Style_Location sl WITH (NOLOCK) on o.StyleUkey = sl.StyleUkey
WHERE o.ID IN (select OrderID from #FinalInspection_Order)
AND  o.ID NOT IN (select OrderID from Order_Location WITH(NOLOCK))

select distinct Location 
from Order_Location with (nolock)
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
    }
}
