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

        public IList<string> GetSewingLineForSetting(string factoryID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@FactoryID", factoryID);

            string sqlGetData = $@"
select id
 from SewingLine sl
where   FactoryID = @FactoryID
        and Junk = 0
order by ID
";
            
            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.AsEnumerable().Select(s => s["id"].ToString()).ToList();
            }
            else
            {
                return new List<string>();
            }
        }
    }
}
