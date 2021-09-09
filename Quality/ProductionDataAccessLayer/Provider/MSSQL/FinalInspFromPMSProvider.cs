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

        public IList<SelectedPO> GetSelectedPOForInspection(List<string> listOrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string whereOrderID = listOrderID.Select(s => $"'{s}'").JoinToString(",");
            string sqlGetData = $@"
select  [OrderID] = o.id,
        o.POID,
        o.StyleID,
        o.SeasonID,
        o.BrandID,
        o.Qty,
        [AvailableQty] = 0,
        [Cartons] = '',
        [Seq] = '',
        [Article] = (SELECT Stuff((select concat( ',',Article)   from Order_Article where ID = o.ID FOR XML PATH('')),1,1,'') )
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
        o.POID,
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
 from PackingList_Detail pld
 where  pld.OrderID in ({whereOrderID}) and
        CTNQty = 1
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
from PackingList_Detail pld
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
from AcceptableQualityLevels
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
		[RowIndex]=ROW_NUMBER() OVER(ORDER BY gdt.id,gdc.id) -1
    from GarmentDefectType gdt with (nolock)
    inner join GarmentDefectCode gdc with (nolock) on gdt.id=gdc.GarmentDefectTypeID
    left join   #FinalInspection_Detail fd on fd.GarmentDefectTypeID = gdt.ID and fd.GarmentDefectCodeID = gdc.ID
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
from DropDownList ddl where
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
    }
}
