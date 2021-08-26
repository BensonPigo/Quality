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
        [Cartons] = ''
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
        o.Qty,
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
        [CTNNo] = CTNStartNo
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
        CTNNo
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
        [CTNNo] = CTNStartNo
from PackingList_Detail pld
left join   #FinalInspection_OrderCarton fc on  fc.OrderID = pld.OrderID and 
                                                fc.PackinglistID = pld.ID and 
                                                fc.CTNNo = pld.CTNStartNo
where   pld.OrderID in (select OrderID from #FinalInspection_Order) and
        pld.CTNQty = 1

";
            return ExecuteList<SelectCarton>(CommandType.Text, sqlGetData, listPar);
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
        [Qty] = isnull(fd.Qty, 0)
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
select  OrderID
into #FinalInspection_Order
from [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where ID = @finalInspectionID

select distinct Article 
from Order_Article with (nolock)
where id in (select OrderID from #FinalInspection_Order)
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

        public List<string> GetSizeList(string finalInspectionID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@finalInspectionID", finalInspectionID);

            string sqlGetMoistureArticleList = @"
select  OrderID
into #FinalInspection_Order
from [ExtendServer].ManufacturingExecution.dbo.FinalInspection_Order with (nolock)
where ID = @finalInspectionID

select distinct SizeCode 
from Order_Qty with (nolock)
where id in (select OrderID from #FinalInspection_Order)
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetMoistureArticleList, listPar);

            if (dtResult.Rows.Count == 0)
            {
                return new List<string>();
            }
            else
            {
                return dtResult.AsEnumerable().Select(s => s["SizeCode"].ToString()).ToList();
            }
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
    }
}
