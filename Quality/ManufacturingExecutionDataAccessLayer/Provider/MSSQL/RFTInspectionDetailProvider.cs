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
            };

            string sqlcmd = @"
select oq.Qty,insp.cnt,* 
from Production.dbo.Order_Qty oq
outer apply(
	select cnt = count(1) 
	from ManufacturingExecution.dbo.Rft_Inspection  i
	where i.OrderId = oq.ID
	and i.Size = oq.SizeCode
)insp
where oq.id = @OrderID and SizeCode = @Size
and oq.Qty - isnull(insp.cnt,0) > 0
";
            DataTable dt = ExecuteDataTable(CommandType.Text, sqlcmd, objParameter);
            return dt;
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

        /*回傳(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳
        /// </summary> 
        /// <param name="Item">成員</param>
        /// <returns>回傳</returns>
        /// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public IList<RFT_Inspection_Detail> Get(RFT_Inspection_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         ID" + Environment.NewLine);
            SbSql.Append("        ,Ukey" + Environment.NewLine);
            SbSql.Append("        ,DefectCode" + Environment.NewLine);
            SbSql.Append("        ,AreaCode" + Environment.NewLine);
            SbSql.Append("        ,Junk" + Environment.NewLine);
            SbSql.Append("        ,PMS_RFTBACriteriaID" + Environment.NewLine);
            SbSql.Append("        ,PMS_RFTRespID" + Environment.NewLine);
            SbSql.Append("        ,GarmentDefectTypeID" + Environment.NewLine);
            SbSql.Append("        ,GarmentDefectCodeID" + Environment.NewLine);
            SbSql.Append("        ,DefectPicture" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("FROM [RFT_Inspection_Detail]" + Environment.NewLine);

            return ExecuteList<RFT_Inspection_Detail>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        /*建立(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
        /// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>

        /*更新(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
        /// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Update(RFT_Inspection_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [RFT_Inspection_Detail]" + Environment.NewLine);
            SbSql.Append("SET" + Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID" + Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID); }
            if (Item.Ukey != null) { SbSql.Append(",Ukey=@Ukey" + Environment.NewLine); objParameter.Add("@Ukey", DbType.String, Item.Ukey); }
            if (Item.DefectCode != null) { SbSql.Append(",DefectCode=@DefectCode" + Environment.NewLine); objParameter.Add("@DefectCode", DbType.String, Item.DefectCode); }
            if (Item.AreaCode != null) { SbSql.Append(",AreaCode=@AreaCode" + Environment.NewLine); objParameter.Add("@AreaCode", DbType.String, Item.AreaCode); }
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk" + Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk); }
            if (Item.PMS_RFTBACriteriaID != null) { SbSql.Append(",PMS_RFTBACriteriaID=@PMS_RFTBACriteriaID" + Environment.NewLine); objParameter.Add("@PMS_RFTBACriteriaID", DbType.String, Item.PMS_RFTBACriteriaID); }
            if (Item.PMS_RFTRespID != null) { SbSql.Append(",PMS_RFTRespID=@PMS_RFTRespID" + Environment.NewLine); objParameter.Add("@PMS_RFTRespID", DbType.String, Item.PMS_RFTRespID); }
            if (Item.GarmentDefectTypeID != null) { SbSql.Append(",GarmentDefectTypeID=@GarmentDefectTypeID" + Environment.NewLine); objParameter.Add("@GarmentDefectTypeID", DbType.String, Item.GarmentDefectTypeID); }
            if (Item.GarmentDefectCodeID != null) { SbSql.Append(",GarmentDefectCodeID=@GarmentDefectCodeID" + Environment.NewLine); objParameter.Add("@GarmentDefectCodeID", DbType.String, Item.GarmentDefectCodeID); }
            if (Item.DefectPicture != null) { SbSql.Append(",DefectPicture=@DefectPicture" + Environment.NewLine); objParameter.Add("@DefectPicture", DbType.String, Item.DefectPicture); }
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate" + Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate); }
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        /*刪除(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
        /// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Delete(RFT_Inspection_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, Item.ID } ,
            };
            SbSql.Append("DELETE FROM [RFT_Inspection_Detail]" + Environment.NewLine);
            SbSql.Append("where 1=1" + Environment.NewLine);
            SbSql.Append("and id = @ID" + Environment.NewLine);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
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
INSERT INTO [RFT_Inspection](
       [OrderID] ,[Article] ,[Location] ,[Size] ,[Line] ,[FactoryID] ,[StyleUkey] ,[FixType]
      ,[ReworkCardNo] ,[Status] ,[AddDate] ,[AddName] ,[ReworkCardType] ,[InspectionDate])
values(
     @OrderID, @Article,@Location,@Size,@Line,@FactoryID,@StyleUkey
    ,@FixType,@ReworkCardNo,@Status, GetDate(),@UserName,@ReworkCardType,@InspectionDate
)

select @@IDENTITY as ID
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
    {picColumn}
    ,[AddDate])
values(
    @@IDENTITY
    ,@DefectCode{detailcnt}
    ,@AreaCode{detailcnt}
    ,0
    ,@PMS_RFTBACriteriaID{detailcnt}
    ,@PMS_RFTRespID{detailcnt}
    ,@GarmentDefectTypeID{detailcnt}
    ,@GarmentDefectCodeID{detailcnt}
    {picValue}
    ,GetDate())
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
INSERT INTO [RFT_Inspection_Detail](
     [ID]
    ,[DefectCode]
    ,[AreaCode]
    ,[Junk]
    ,[PMS_RFTBACriteriaID]
    ,[PMS_RFTRespID]
    ,[GarmentDefectTypeID]
    ,[GarmentDefectCodeID]
    ,DefectPicture
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
    ,@DefectPicture
    ,GetDate())
";
           
            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }

        #endregion
    }
}
