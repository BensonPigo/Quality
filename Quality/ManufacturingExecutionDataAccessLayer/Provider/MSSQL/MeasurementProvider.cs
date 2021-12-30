using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using System.Windows.Documents;
using DatabaseObject.RequestModel;
using DatabaseObject.ProductionDB;
using System.Linq;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class MeasurementProvider : SQLDAL, IMeasurementProvider
    {
        #region 底層連線
        public MeasurementProvider(string conString) : base(conString) { }
        public MeasurementProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
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
        public IList<Measurement> Get(Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,Tol1"+ Environment.NewLine);
            SbSql.Append("        ,Tol2"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Code"+ Environment.NewLine);
            SbSql.Append("        ,SizeCode"+ Environment.NewLine);
            SbSql.Append("        ,SizeSpec"+ Environment.NewLine);
            SbSql.Append("        ,Ukey"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,SizeGroup"+ Environment.NewLine);
            SbSql.Append("        ,MeasurementTranslateUkey"+ Environment.NewLine);
            SbSql.Append("FROM [Measurement] WITH(NOLOCK)" + Environment.NewLine);



            return ExecuteList<Measurement>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Measurement]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,Tol1"+ Environment.NewLine);
            SbSql.Append("        ,Tol2"+ Environment.NewLine);
            SbSql.Append("        ,Description"+ Environment.NewLine);
            SbSql.Append("        ,Code"+ Environment.NewLine);
            SbSql.Append("        ,SizeCode"+ Environment.NewLine);
            SbSql.Append("        ,SizeSpec"+ Environment.NewLine);
            SbSql.Append("        ,Ukey"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,Junk"+ Environment.NewLine);
            SbSql.Append("        ,SizeGroup"+ Environment.NewLine);
            SbSql.Append("        ,MeasurementTranslateUkey"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @StyleUkey"); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);
            SbSql.Append("        ,@Tol1"); objParameter.Add("@Tol1", DbType.String, Item.Tol1);
            SbSql.Append("        ,@Tol2"); objParameter.Add("@Tol2", DbType.String, Item.Tol2);
            SbSql.Append("        ,@Description"); objParameter.Add("@Description", DbType.String, Item.Description);
            SbSql.Append("        ,@Code"); objParameter.Add("@Code", DbType.String, Item.Code);
            SbSql.Append("        ,@SizeCode"); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);
            SbSql.Append("        ,@SizeSpec"); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);
            SbSql.Append("        ,@Ukey"); objParameter.Add("@Ukey", DbType.String, Item.Ukey);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
            SbSql.Append("        ,@SizeGroup"); objParameter.Add("@SizeGroup", DbType.String, Item.SizeGroup);
            SbSql.Append("        ,@MeasurementTranslateUkey"); objParameter.Add("@MeasurementTranslateUkey", DbType.String, Item.MeasurementTranslateUkey);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
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
        public int Update(Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Measurement]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.StyleUkey != null) { SbSql.Append("StyleUkey=@StyleUkey"+ Environment.NewLine); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);}
            if (Item.Tol1 != null) { SbSql.Append(",Tol1=@Tol1"+ Environment.NewLine); objParameter.Add("@Tol1", DbType.String, Item.Tol1);}
            if (Item.Tol2 != null) { SbSql.Append(",Tol2=@Tol2"+ Environment.NewLine); objParameter.Add("@Tol2", DbType.String, Item.Tol2);}
            if (Item.Description != null) { SbSql.Append(",Description=@Description"+ Environment.NewLine); objParameter.Add("@Description", DbType.String, Item.Description);}
            if (Item.Code != null) { SbSql.Append(",Code=@Code"+ Environment.NewLine); objParameter.Add("@Code", DbType.String, Item.Code);}
            if (Item.SizeCode != null) { SbSql.Append(",SizeCode=@SizeCode"+ Environment.NewLine); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);}
            if (Item.SizeSpec != null) { SbSql.Append(",SizeSpec=@SizeSpec"+ Environment.NewLine); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);}
            if (Item.Ukey != null) { SbSql.Append(",Ukey=@Ukey"+ Environment.NewLine); objParameter.Add("@Ukey", DbType.String, Item.Ukey);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk"+ Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk);}
            if (Item.SizeGroup != null) { SbSql.Append(",SizeGroup=@SizeGroup"+ Environment.NewLine); objParameter.Add("@SizeGroup", DbType.String, Item.SizeGroup);}
            if (Item.MeasurementTranslateUkey != null) { SbSql.Append(",MeasurementTranslateUkey=@MeasurementTranslateUkey"+ Environment.NewLine); objParameter.Add("@MeasurementTranslateUkey", DbType.String, Item.MeasurementTranslateUkey);}
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
        public int Delete(Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Measurement]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<Order_Qty> GetAtricle(string OrderID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, OrderID } ,
            };
            string sqlcmd = @"
Select distinct Article
From MainServer.Production.dbo.Order_Qty WITH(NOLOCK)
where ID = @OrderID
";
             return ExecuteList<Order_Qty>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<Measurement_Request> Get_OrdersPara(string OrderID, string FactoryID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, OrderID } ,
                { "@FactoryID", DbType.String, FactoryID } ,
            };
            string sqlcmd = @"
select o.OrderTypeID,[OrderID] = o.ID,o.StyleID,o.SeasonID
,[Unit] = IIF(isnull(o.sizeUnit,'') = '',s.SizeUnit,o.SizeUnit)
,[Factory] = o.FtyGroup--o.FactoryID
from MainServer.Production.dbo.Orders o WITH(NOLOCK)
left join MainServer.Production.dbo.Style s WITH(NOLOCK) on o.StyleUkey = s.Ukey
where o.ID = @OrderID
--and o.FactoryID = @FactoryID
";
            return ExecuteList<Measurement_Request>(CommandType.Text, sqlcmd, objParameter);
        }

        public int Get_Total_Measured_Qty()
        {
            int rtQty = 0;
            string sqlcmd = @"
select [ttlCnt] = count (1)
from (
    select distinct styleukey, sizecode, no
    from RFT_Inspection_Measurement WITH(NOLOCK)
    where format (adddate, 'yyyy/MM/dd') = format (getdate(), 'yyyy/MM/dd')
) rim
";
            DataTable dt = ExecuteDataTable(CommandType.Text, sqlcmd, new SQLParameterCollection());
            if (dt != null)
            {
                rtQty = Convert.ToInt32(dt.Rows[0]["ttlCnt"]);
            }

            return rtQty;
        }

        public int Get_Measured_Qty(Measurement_Request measurement)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, measurement.OrderID } ,
                { "@Article", DbType.String, measurement.Article } ,
            };

            int rtQty = 0;
            string sqlcmd = @"
-- 等於 Find 後下方 Diff 出現的次數
select [ttlCnt] = count (1) 
from (
    select distinct styleukey, sizecode, no
    from RFT_Inspection_Measurement WITH(NOLOCK)
    where OrderID = @OrderID
        and Article = @Article
) rim
";
            DataTable dt = ExecuteDataTable(CommandType.Text, sqlcmd, objParameter);
            if (dt != null)
            {
                rtQty = Convert.ToInt32(dt.Rows[0]["ttlCnt"]);
            }

            return rtQty;
        }

        public DataTable Get_CalculateSizeSpec(string diffValue, string Tol)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@DiffValue", DbType.String, diffValue } ,
                { "@Tol", DbType.String, Tol } ,
            };

            string sqlcmd = @"
select value = dbo.calculateSizeSpec(@DiffValue, @Tol,'INCH');
";
            DataTable dt = ExecuteDataTable(CommandType.Text, sqlcmd, objParameter);
            return dt;
        }

        public DataTable Get_Measured_Detail(Measurement_Request measurement)
        {
            string styleUkey = string.Empty;
            DataTable dtStyle = ExecuteDataTable(CommandType.Text, $@"Select StyleUkey from MainServer.Production.dbo.Orders WITH(NOLOCK) where id = '{measurement.OrderID}'", new SQLParameterCollection());
            if (dtStyle.Rows.Count > 0)
            {
                styleUkey = dtStyle.Rows[0]["StyleUkey"].ToString();
            }

            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@_OrderID", DbType.String, measurement.OrderID } ,
                { "@_StyleUkey", DbType.String, styleUkey } ,
                //{ "@Team", DbType.String, measurement.Team } ,
                //{ "@Line", DbType.String, measurement.Line } ,
                { "@_Factory", DbType.String, measurement.Factory } ,
                { "@_TypeUnit", DbType.String, measurement.Unit } ,
            };


            string sqlcmd = @"

declare @styleukey nvarchar(10) = @_StyleUkey 
--declare @team nvarchar(1) = @Team
--declare @line nvarchar(3) = @Line
declare @factory nvarchar(3) = @_Factory
declare @typeUnit nvarchar(5)= @_TypeUnit
declare @ex nvarchar(max)=''
declare @ex2 nvarchar(max)=''
declare @col nvarchar(max)=''
declare @size nvarchar(max)
declare @SizeCode nvarchar(8)
declare @OldSizeCode nvarchar(8)=''
declare @no nvarchar(66)
declare @time nvarchar(5)
declare @diffno int='1'
declare @orderid nvarchar(16) = @_OrderID
declare @MDivision nvarchar(5) = (select MDivisionID from MainServer.Production.dbo.Factory WITH(NOLOCK) where id = @Factory)
--declare @shiftTabele table(MDivision varchar(8),Shift varchar(5),StartDate date,BeginTime time,EndTime time,ActualBeginTime datetime,ActualEndTime datetime)
--declare @workStartDatetime datetime
--declare @workEndDatetime datetime

--
--INSERT INTO @shiftTabele
--SELECT * FROM [dbo].[GetWorkShiftTable](@MDivision,GETDATE(),@factory)

--SELECT  @workStartDatetime=ActualBeginTime,@workEndDatetime=ActualEndTime
--FROM @shiftTabele

select *
into #tmp_Inspection_Measurement
from RFT_Inspection_Measurement im WITH(NOLOCK)
where im.StyleUkey = @styleukey 
--and (@workStartDatetime <= im.AddDate AND im.AddDate <= @workEndDatetime)
--and im.Team = @team
--and im.Line = @line
and im.FactoryID = @factory 
and im.OrderID = @orderid 


select m.Ukey
	,Description= iif(isnull(b.DescEN,'') = '',m.Description,b.DescEN)
	,m.Tol1
	,m.Tol2
	,m.Code
	,m.SizeCode 
	,[MeasurementSizeSpec] = m.SizeSpec 
	,[InspectionMeasurementSizeSpec] = im.SizeSpec
	,[diff]= max(dbo.calculateSizeSpec(m.SizeSpec,im.SizeSpec, @typeUnit))
	,im.AddDate
	,[HeadSizeCode] = FORMAT(im.AddDate,'yyyy/MM/dd HH:mm:ss')
into #tmp 
from Measurement m with(nolock)
inner join (select distinct StyleUkey, SizeCode from #tmp_Inspection_Measurement) t on m.StyleUkey = t.StyleUkey and m.SizeCode = t.SizeCode
left join #tmp_Inspection_Measurement im on im.MeasurementUkey = m.Ukey 
LEFT JOIN [ManufacturingExecution].[dbo].[MeasurementTranslate] b ON  m.MeasurementTranslateUkey = b.UKey
where m.junk = 0
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

        public IList<Measurement> GetMeasurementsByPOID(string CustPONO, string userID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@CustPONO", DbType.String, CustPONO },
                { "@userID", DbType.String, userID },
            };


            string sqlcmd = @"
declare @SizeUnit varchar(8)
declare @StyleUkey bigint

select  @StyleUkey = StyleUkey
from    SciProduction_Orders with (nolock)
where   ID IN (
    select POID
    from SciProduction_Orders WITH(NOLOCK)
    where CustPONO = @CustPONO
)

exec CopyStyle_ToMeasurement @userID,@StyleUkey;

select  @SizeUnit = SizeUnit
from    MainServer.Production.dbo.Style WITH(NOLOCK)
where   Ukey = @StyleUkey

SELECT  a.StyleUkey
        ,a.Tol1
        ,a.Tol2
        ,a.Description
        ,a.Code
        ,a.SizeCode
        ,a.SizeSpec
        ,a.Ukey
        ,a.AddDate
        ,a.AddName
        ,a.Junk
        ,a.SizeGroup
        ,a.MeasurementTranslateUkey
        , [IsPatternMeas] = case when a.Description like '%pattern measn%'  then convert(bit, 1)
			                when  isnull(a.Tol1, '0') = '0' and isnull(a.Tol2, '0') = '0' and isnull(a.SizeSpec, '0') = '0' then convert(bit, 1)
			                when  isnull(a.SizeCode,'') = '' then convert(bit, 1)
			                when  a.SizeSpec like '%[a-zA-Z]%' then convert(bit, 1)
			                when  (UPPER(@SizeUnit) = 'INCH' and  a.SizeSpec like '%.%') then convert(bit, 1)
			                when  (UPPER(@SizeUnit) = 'CM' and  a.SizeSpec like '%/%') then convert(bit, 1)
			                else  convert(bit, 0) end
FROM [ManufacturingExecution].[dbo].[Measurement] a with(nolock)
where a.junk=0 
and a.StyleUkey = @StyleUkey 

";

            return ExecuteList<Measurement>(CommandType.Text, sqlcmd, objParameter, 80);
        }
        #endregion
    }
}
