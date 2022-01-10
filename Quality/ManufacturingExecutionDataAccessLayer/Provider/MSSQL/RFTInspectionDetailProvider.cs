using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class RFTInspectionDetailProvider : SQLDAL, IRFTInspectionDetailProvider
    {
        #region 底層連線
        public RFTInspectionDetailProvider(string conString) : base(conString) { }
        public RFTInspectionDetailProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        public DataTable ChkInspQty(RFT_Inspection filter)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, filter.OrderID } ,
                { "@Size", DbType.String, filter.Size } ,
                { "@Article", DbType.String, filter.Article},
            };

            string sqlcmd = @"
select oq.Qty,insp.cnt,* 
from MainServer.Production.dbo.Order_Qty oq WITH(NOLOCK)
outer apply(
	select cnt = count(1) 
	from ManufacturingExecution.dbo.Rft_Inspection  i WITH(NOLOCK)
	where i.OrderId = oq.ID
	and i.Size = oq.SizeCode
    and i.Article = oq.Article
    and Status <> 'Dispose'
)insp
where oq.id = @OrderID and SizeCode = @Size and Article = @Article
and oq.Qty - isnull(insp.cnt,0) > 0
";         
            return ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<RFT_Inspection_Detail> Top3Defects(RFT_Inspection Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@FactoryID", DbType.String, Item.FactoryID } ,
                { "@Line", DbType.String, Item.Line } ,
                { "@InspectionDate", DbType.DateTime, Item.InspectionDate } ,
            };

            SbSql.Append(
                @"
;with DefectCntDt as (
    select top 3 [DefectCode] = gdc.Description
        ,id.GarmentDefectCodeID
        ,[DefectCnt] = count(*) 
    from RFT_Inspection i WITH (NOLOCK)
    inner join RFT_Inspection_Detail id  WITH (NOLOCK) on id.ID = i.ID
    left join [dbo].[SciProduction_GarmentDefectCode] gdc with (nolock) on id.GarmentDefectCodeID = gdc.id 
    where 1=1
  And ((i.AddDate >= @InspectionDate and i.AddDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate))) 
  or (i.EditDate >= @InspectionDate and i.EditDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate)))) 
    and i.Line = @Line
    and i.FactoryID = @FactoryID
    and id.junk = 0
    group by gdc.Description,id.GarmentDefectCodeID 
    order by count(*) desc,gdc.Description asc
)

select [DefectCode] = replace(DefectCode,'/','/ '),
       [AreaCode] = (select Stuff((
                    select top 3 concat( '/ ',a.Code)   
                    from RFT_Inspection i WITH (NOLOCK)
                    inner join RFT_Inspection_Detail id  WITH (NOLOCK) on id.ID = i.ID
                    inner join dbo.Area a  WITH (NOLOCK) on id.AreaCode = a.Code
                    where 1=1
                        And ((i.AddDate >= @InspectionDate and i.AddDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate))) 
  or (i.EditDate >= @InspectionDate and i.EditDate <= DATEADD(SECOND, -1, DATEADD(day, 1,@InspectionDate)))) 
                    and i.Line = @Line
                    and i.FactoryID = @FactoryID        
                    and id.GarmentDefectCodeID = DefectCntDt.GarmentDefectCodeID 
                    and id.junk = 0
                    group by a.Code  order by count(*) desc  FOR XML PATH('')),1,1,'') )
		,GarmentDefectCodeID
from DefectCntDt
" + Environment.NewLine);
        

            return ExecuteList<RFT_Inspection_Detail>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Create_Master_Detail(RFT_Inspection Master, List<RFT_Inspection_Detail> Detail)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, string.IsNullOrEmpty(Master.OrderID)?"" : Master.OrderID } ,
                { "@Article", DbType.String, string.IsNullOrEmpty(Master.Article)?"" : Master.Article } ,
                { "@Location", DbType.String, string.IsNullOrEmpty(Master.Location)?"" : Master.Location } ,
                { "@Size", DbType.String, string.IsNullOrEmpty(Master.Size)?"" : Master.Size } ,
                { "@Line", DbType.String, string.IsNullOrEmpty(Master.Line)?"" : Master.Line } ,
                { "@FactoryID", DbType.String, string.IsNullOrEmpty(Master.FactoryID)?"" : Master.FactoryID } ,
                { "@StyleUkey", DbType.String, Master.StyleUkey } ,
                { "@FixType", DbType.String, string.IsNullOrEmpty(Master.FixType)?"" : Master.FixType } ,
                { "@ReworkCardNo", DbType.String,  string.IsNullOrEmpty(Master.ReworkCardNo)?"" : Master.ReworkCardNo } ,
                { "@Status", DbType.String, string.IsNullOrEmpty(Master.Status)?"" : Master.Status } ,
                { "@UserName", DbType.String, string.IsNullOrEmpty(Master.AddName)?"" : Master.AddName } ,
                { "@ReworkCardType", DbType.String, string.IsNullOrEmpty(Master.ReworkCardType)?"" : Master.ReworkCardType } ,
                { "@InspectionDate", DbType.DateTime, Master.InspectionDate } ,
            };

            string sqlcmd = $@"
SET XACT_ABORT ON
declare @ID as bigint

INSERT INTO [RFT_Inspection](
       [OrderID] ,[Article] ,[Location] ,[Size] ,[Line] ,[FactoryID] ,[StyleUkey] ,[FixType]
      ,[ReworkCardNo] ,[Status] ,[AddDate] ,[AddName] ,[ReworkCardType] ,[InspectionDate])
values(
     @OrderID, @Article,@Location,@Size,@Line,@FactoryID,@StyleUkey
    ,@FixType,@ReworkCardNo,@Status, GetDate(),@UserName,@ReworkCardType,@InspectionDate
)

select @ID = @@IDENTITY
";

            int detailcnt = 1;
            foreach (var item in Detail)
            {
                string picColumn = string.Empty;
                string picValue = string.Empty;

                objParameter.Add($"@DefectCode{detailcnt}", string.IsNullOrEmpty(item.DefectCode) ? "" : item.DefectCode);
                objParameter.Add($"@AreaCode{detailcnt}", string.IsNullOrEmpty(item.AreaCode) ? "" : item.AreaCode);
                objParameter.Add($"@PMS_RFTBACriteriaID{detailcnt}", string.IsNullOrEmpty(item.PMS_RFTBACriteriaID) ? "" : item.PMS_RFTBACriteriaID);
                objParameter.Add($"@PMS_RFTRespID{detailcnt}", string.IsNullOrEmpty(item.PMS_RFTRespID) ? "" : item.PMS_RFTRespID);
                objParameter.Add($"@GarmentDefectTypeID{detailcnt}", string.IsNullOrEmpty(item.GarmentDefectTypeID) ? "" : item.GarmentDefectTypeID);
                objParameter.Add($"@GarmentDefectCodeID{detailcnt}", string.IsNullOrEmpty(item.GarmentDefectCodeID) ? "" : item.GarmentDefectCodeID);
                objParameter.Add($"@DefectPicture{detailcnt}", item.DefectPicture);

                if (item.DefectPicture != null)
                {
                    picColumn = ",[DefectPicture]";
                    picValue = $",@DefectPicture{detailcnt}";
                }

                sqlcmd += $@"
INSERT INTO [RFT_Inspection_Detail](
     [ID]
    ,[DefectCode]
    ,[AreaCode]
    ,[Junk]
    ,[PMS_RFTBACriteriaID]
    ,[PMS_RFTRespID]
    ,[GarmentDefectTypeID]
    ,[GarmentDefectCodeID]
                                /*2022/01/10 PMSFile上線，因此去掉Image寫入原本DB的部分*/
    ,[AddDate]) 
values(
    @ID
    ,@DefectCode{detailcnt}
    ,@AreaCode{detailcnt}
    ,0
    ,@PMS_RFTBACriteriaID{detailcnt}
    ,@PMS_RFTRespID{detailcnt}
    ,@GarmentDefectTypeID{detailcnt}
    ,@GarmentDefectCodeID{detailcnt}
    
    ,GetDate())

INSERT INTO PMSFile.dbo.[RFT_Inspection_Detail](
     [ID],Ukey
    {picColumn}) 
values(
    @ID,(SELECT Max(Ukey) + 1 from RFT_Inspection_Detail where id=@ID)
    {picValue})
";
                detailcnt++;
            }

            // update Rework Card Sataus = 'Rework'（代表 Using 使用中）
            sqlcmd += @"
update ReworkCard 
set Status = 'Rework'
,EditName = @UserName
,EditDate = GETDATE()
where  No = @ReworkCardNo 
and  Type = @FixType
and  Line = @Line
and  FactoryID = @FactoryID
";

            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }

        public int Create_Detail(RFT_Inspection_Detail Detail)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add($"@ID", string.IsNullOrEmpty(Detail.ID.ToString()) ? "" : Detail.ID.ToString());
            objParameter.Add($"@DefectCode", string.IsNullOrEmpty(Detail.DefectCode) ? "" : Detail.DefectCode);
            objParameter.Add($"@AreaCode", string.IsNullOrEmpty(Detail.AreaCode) ? "" : Detail.AreaCode);
            objParameter.Add($"@PMS_RFTBACriteriaID", string.IsNullOrEmpty(Detail.PMS_RFTBACriteriaID) ? "" : Detail.PMS_RFTBACriteriaID);
            objParameter.Add($"@Junk", Detail.Junk);
            objParameter.Add($"@PMS_RFTRespID", string.IsNullOrEmpty(Detail.PMS_RFTRespID) ? "" : Detail.PMS_RFTRespID);
            objParameter.Add($"@GarmentDefectTypeID", string.IsNullOrEmpty(Detail.GarmentDefectTypeID) ? "" : Detail.GarmentDefectTypeID);
            objParameter.Add($"@GarmentDefectCodeID", string.IsNullOrEmpty(Detail.GarmentDefectCodeID) ? "" : Detail.GarmentDefectCodeID);
            objParameter.Add($"@DefectPicture", Detail.DefectPicture);

            string sqlcmd = string.Empty;

            if (Detail.DefectPicture != null) { objParameter.Add("@DefectPicture", Detail.DefectPicture); }
            else { objParameter.Add("@DefectPicture", System.Data.SqlTypes.SqlBinary.Null); }

            sqlcmd += $@"
SET XACT_ABORT ON
INSERT INTO [RFT_Inspection_Detail](
     [ID]
    ,[DefectCode]
    ,[AreaCode]
    ,[Junk]
    ,[PMS_RFTBACriteriaID]
    ,[PMS_RFTRespID]
    ,[GarmentDefectTypeID]
    ,[GarmentDefectCodeID]
                        ----2022/01/10 PMSFile上線，因此去掉Image寫入原本DB的部分
    ,[AddDate])
values(
     @ID
    ,@DefectCode
    ,@AreaCode
    ,@Junk
    ,@PMS_RFTBACriteriaID
    ,@PMS_RFTRespID
    ,@GarmentDefectTypeID
    ,@GarmentDefectCodeID

    ,GetDate())

INSERT INTO PMSFile.dbo.[RFT_Inspection_Detail](
     [ID],Ukey,DefectPicture) 
values(
    @ID,(SELECT Max(Ukey) + 1 from RFT_Inspection_Detail where id=@ID)
    ,@DefectPicture
)
";
           
            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }

        #endregion
    }
}
