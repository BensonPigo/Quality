using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System.Linq;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentTestDetailProvider : SQLDAL,IGarmentTestDetailProvider
    {
        #region 底層連線
        public GarmentTestDetailProvider(string ConString) : base(ConString) { }
        public GarmentTestDetailProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public IList<Order_Qty> GetSizeCode(string OrderID, string Article)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, OrderID } ,
                { "@Article", DbType.String, Article } ,
            };

            string sqlcmd = string.Empty;
            if (string.IsNullOrEmpty(OrderID))
            {
                sqlcmd = "select distinct SizeCode from Style_SizeCode WITH(NOLOCK) order by SizeCode";
            }
            else
            {
                sqlcmd = @"
select distinct SizeCode from Order_Qty WITH(NOLOCK)
where ID = @OrderID and Article = @Article 
order by SizeCode";
            }

            return ExecuteList<Order_Qty>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<Order_Qty> GetSizeCode(string StyleID, string SeasonID, string BrandID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@StyleID", DbType.String, StyleID } ,
                { "@SeasonID", DbType.String, SeasonID } ,
                { "@BrandID", DbType.String, BrandID } ,
            };

            string sqlcmd = @"
select distinct SizeCode 
from Style s WITH(NOLOCK)
inner join Style_SizeCode sc WITH(NOLOCK) on s.Ukey = sc.StyleUkey
where s.ID = @StyleID and s.SeasonID = @SeasonID and s.BrandID = @BrandID
";

            return ExecuteList<Order_Qty>(CommandType.Text, sqlcmd, objParameter);
        }

        public List<string> GetScales()
        {
            string sqlcmd = @"
select ID = ''
union all
select ID from Scale WITH(NOLOCK)  WHERE Junk=0 
order by ID
";
            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, new SQLParameterCollection());

            return dt.Rows.OfType<DataRow>().Select(dr => dr.Field<string>("ID")).ToList();
        }

        public IList<GarmentTest_Detail_ViewModel> Get_GarmentTestDetail(GarmentTest filter)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, filter.ID } ,
            };
            string sqlcmd = @"
select 
gd.No
,gd.ReportNo
,gd.ID
,gd.OrderID
,gd.SizeCode
,gd.inspdate
,gd.MtlTypeID
,[Result] = case when gd.Result = 'P' then 'Pass' 
				 when gd.Result = 'F' then 'Fail' 
				 else '' end
,gd.NonSeamBreakageTest
,[SeamBreakageResult] = case when gd.SeamBreakageResult = 'P' then 'Pass' 
							 when gd.SeamBreakageResult = 'F' then 'Fail' 
							 else '' end
,[OdourResult] = case when gd.OdourResult = 'P' then 'Pass' 
					  when gd.OdourResult = 'F' then 'Fail' 
					  else '' end
,[WashResult] = case when gd.WashResult = 'P' then 'Pass' 
					 when gd.WashResult = 'F' then 'Fail' 
					 else '' end
,gd.inspector
,[GarmentTest_Detail_Inspector] = isnull(InspectorName.Name_Extno,'')
,gd.Remark
,gd.Sender
,gd.SendDate
,gd.Receiver
,gd.ReceiveDate
,[GarmentTest_Detail_AddName] = CONCAT(gd.AddName,'-',CreatBy.Name,'',gd.AddDate)
,[GarmentTest_Detail_EditName] = CONCAT(gd.EditName,'-',EditBy.Name,'',gd.EditDate)
,gd.SubmitDate,gd.ArrivedQty,gd.LineDry,gd.Temperature,gd.TumbleDry,gd.Machine,gd.HandWash
,gd.Composition,gd.Neck,gd.Status,gd.LOtoFactory,gd.MtlTypeID,gd.Above50NaturalFibres,gd.Above50SyntheticFibres
,gdi.TestAfterPicture,gdi.TestBeforePicture
from GarmentTest_Detail gd WITH(NOLOCK)
left join [ExtendServer].PMSFile.dbo.GarmentTest_Detail gdi WITH(NOLOCK) on gd.ID = gdi.ID AND gd.No = gdi.No
left join Pass1 CreatBy WITH(NOLOCK) on CreatBy.ID = gd.AddName
left join Pass1 EditBy WITH(NOLOCK) on EditBy.ID = gd.EditName
outer apply(
	select Name_Extno 
	from View_ShowName
	where ID=gd.inspector
)InspectorName
where gd.ID = @ID
" + Environment.NewLine;
            

            return ExecuteList<GarmentTest_Detail_ViewModel>(CommandType.Text, sqlcmd, objParameter);
        }

        public bool Recalculate_Result()
        {
            return true;
        }

        public int Delete_GarmentTestDetail(string ID, string No)
        {  
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };

            string sqlcmd = @"
SET XACT_ABORT ON

Delete GarmentTest_Detail_Shrinkage  where id = @ID and NO = @No
Delete GarmentTest_Detail_Apperance where id = @ID and NO = @No
Delete GarmentTest_Detail_FGWT where id = @ID and NO = @No
Delete GarmentTest_Detail_FGPT where id = @ID and NO = @No
Delete Garment_Detail_Spirality where id = @ID and NO = @No

Delete GarmentTest_Detail where id = @ID and NO = @No
Delete [ExtendServer].PMSFile.dbo.GarmentTest_Detail where id = @ID and NO = @No
";
            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }

        public bool Update_Sender(string ID, string No, string UserID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
                { "@UserID", DbType.String, UserID } ,
            };

            string sqlcmd = @"
update GarmentTest_Detail
set Sender = @UserID , SendDate = GETDATE()
where ID = @ID and No = @No
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public bool Update_Receive(string ID, string No, string UserID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
                { "@UserID", DbType.String, UserID } ,
            };

            string sqlcmd = @"
update GarmentTest_Detail
set Receiver = @UserID , ReceiveDate = GETDATE()
where ID = @ID and No = @No

";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public bool Chk_AllResult(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };

            string sqlcmd = @"

-- FGPT
select distinct [Result] = CASE WHEN  t.TestUnit = 'N' AND t.[TestResult] !='' THEN IIF( Cast( t.[TestResult] as float) >= cast( t.Criteria as float) ,'Pass' ,'Fail')
  WHEN  t.TestUnit = 'mm' THEN IIF(  t.[TestResult] = '<=4' OR t.[TestResult] = '≦4','Pass' , IIF( t.[TestResult]='>4','Fail','')  )
  WHEN  t.TestUnit = 'Pass/Fail'  THEN t.[TestResult]
   ELSE ''
END
from GarmentTest_Detail_FGPT t WITH(NOLOCK)
where t.ID = @ID
and t.No = @No

union

--FGWT
select distinct
[Result] = IIF(Scale IS NOT NULL
    ,IIF(Scale='4-5' OR Scale ='5','Pass',IIF(Scale='','','Fail'))
    ,IIF( (BeforeWash IS NOT NULL AND AfterWash IS NOT NULL AND Criteria IS NOT NULL AND Shrinkage IS NOT NULL)
          or (Type = 'spirality: Garment - in percentage (average)')
          or (Type = 'spirality: Garment - in percentage (average) (Top Method A)')
          or (Type = 'spirality: Garment - in percentage (average) (Top Method B)')
          or (Type = 'spirality: Garment - in percentage (average) (Bottom Method A)')
          or (Type = 'spirality: Garment - in percentage (average) (Bottom Method B)')
   ,( IIF( TestDetail = '%' OR TestDetail = 'Range%'   
   -- % 為ISP20201331舊資料、Range% 為ISP20201606加上的新資料，兩者都視作百分比
      ---- 百分比 判斷方式
      ,IIF( ISNULL(Criteria,0)  <= ISNULL(Shrinkage,0) AND ISNULL(Shrinkage,0) <= ISNULL(Criteria2,0)
       , 'Pass'
       , 'Fail'
      )
      ---- 非百分比 判斷方式
      ,IIF( ISNULL(AfterWash,0) - ISNULL(BeforeWash,0) <= ISNULL(Criteria,0)
       ,'Pass'
       ,'Fail'
      )
    )
   )
   ,''
 )
)
from GarmentTest_Detail_FGWT f WITH(NOLOCK)
where ID = @ID
and No = @No

--
";
            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);
            if (dt != null && dt.Rows.Count > 0)
            {
                if (dt.AsEnumerable().Where(x => x["Result"].Equals("Fail")).Any())
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        public string Get_LastResult(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            string sqlcmd = @"
select Result 
from  GarmentTest_Detail t WITH(NOLOCK)
where No = (
	select Max(No) from GarmentTest_Detail s WITH(NOLOCK)
	where s.ID = t.ID
)
and ID = @ID
";
            return Convert.ToString(ExecuteScalar(CommandType.Text, sqlcmd, objParameter));
        }

        public bool Update_GarmentTestDetail(GarmentTest_Detail_ViewModel source)
        {
            string LOtoFactory = string.IsNullOrEmpty(source.LOtoFactory) ? string.Empty : source.LOtoFactory;
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", source.ID } ,
                { "@No", source.No } ,
                { "@SubmitDate", DbType.Date, source.SubmitDate},
                { "@ArrivedQty", source.ArrivedQty } ,
                { "@LOtoFactory", LOtoFactory} ,
                { "@Remark", source.Remark ?? ""} ,
                { "@LineDry", DbType.Boolean, source.LineDry } ,
                { "@Temperature", source.Temperature } ,
                { "@TumbleDry", DbType.Boolean, source.TumbleDry } ,
                { "@Machine", source.Machine } ,
                { "@HandWash",DbType.Boolean, source.HandWash } ,
                { "@Composition", source.Composition } ,
                { "@Neck", DbType.Boolean, source.Neck } ,
                { "@Above50NaturalFibres", DbType.Boolean, source.Above50NaturalFibres } ,
                { "@Above50SyntheticFibres", DbType.Boolean, source.Above50SyntheticFibres } ,
                { "@EditName", source.EditName } ,
                { $"@NonSeamBreakageTest",DbType.Boolean,  source.NonSeamBreakageTest},
            };

            if (source.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", source.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }

            if (source.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", source.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            string sqlcmd = $@"
SET XACT_ABORT ON

update GarmentTest_Detail set
    SubmitDate = @SubmitDate,
    ArrivedQty =  @ArrivedQty,
    LOtoFactory =  @LOtoFactory,
    Remark =  @Remark,
    LineDry =  @LineDry,
    Temperature =  @Temperature,
    TumbleDry =  @TumbleDry,
    Machine =  @Machine,
    HandWash =  @HandWash,
    Composition =  @Composition,
    Neck = @Neck ,
    Above50NaturalFibres =  @Above50NaturalFibres,
    Above50SyntheticFibres =  @Above50SyntheticFibres,
    NonSeamBreakageTest = @NonSeamBreakageTest,
    EditName = @EditName,
    EditDate = GetDate()
where ID = @ID and No = @No



if not exists (select 1 from [ExtendServer].PMSFile.dbo.GarmentTest_Detail where ID = @ID and No = @No)
begin
    INSERT INTO [ExtendServer].PMSFile.dbo.GarmentTest_Detail (ID,No,TestBeforePicture,TestAfterPicture)
    VALUES (@ID,@No,@TestBeforePicture,@TestAfterPicture)
end
else
begin
    update [ExtendServer].PMSFile.dbo.GarmentTest_Detail set
        TestBeforePicture = @TestBeforePicture,
        TestAfterPicture = @TestAfterPicture
    where ID = @ID and No = @No
end
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public bool Save_Detail_Picture(GarmentTest_Detail source)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
             {
                { "@ID", source.ID } ,
                { "@No", source.No } ,
            }; 

            if (source.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", source.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }

            if (source.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", source.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            string sqlcmd = $@"
SET XACT_ABORT ON

/* 2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
update GarmentTest_Detail 
set
    TestBeforePicture = @TestBeforePicture,
    TestAfterPicture = @TestAfterPicture
where ID = @ID and No = @No
*/

if not exists (select 1 from [ExtendServer].PMSFile.dbo.GarmentTest_Detail where ID = @ID and No = @No)
begin
    INSERT INTO [ExtendServer].PMSFile.dbo.GarmentTest_Detail (ID,No,TestBeforePicture,TestAfterPicture)
    VALUES (@ID,@No,@TestBeforePicture,@TestAfterPicture)
end
else
begin
    update [ExtendServer].PMSFile.dbo.GarmentTest_Detail set
        TestBeforePicture = @TestBeforePicture,
        TestAfterPicture = @TestAfterPicture
    where ID = @ID and No = @No
end
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public bool Encode_GarmentTestDetail(string ID, string No, string Status)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
                { "@Status", DbType.String, Status } ,
            };
            if (Status == "Confirmed")
            {
                objParameter.Add("@InspDate", DbType.DateTime, DateTime.Now);
            }
            else
            {
                objParameter.Add("@InspDate", DbType.DateTime, DBNull.Value);
            }

            string sqlcmd = @"
Update GarmentTest_Detail set Status = @Status, InspDate = @InspDate  where id = @ID and No = @No
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public DataTable Get_Mail_Content(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };

            string sqlcmd = @"
select  
    g.StyleID
    ,g.BrandID
    ,g.SeasonID
    ,g.Article
    ,g.OrderID
    ,[SpecialMark] = SpecialMark.Value
    ,gd.No
    ,gd.SizeCode
    ,[TestDate] = format(gd.inspdate, 'yyyy/MM/dd')
    ,[Result] = case when g.Result = 'P' then 'Pass' when g.Result = 'F' then 'Fail' else '' end
    ,[450 Result] = case when gd.SeamBreakageResult = 'P' then 'Pass' when gd.SeamBreakageResult = 'F' then 'Fail' else '' end
    ,[451 Result] = case when gd.OdourResult = 'P' then 'Pass' when gd.OdourResult = 'F' then 'Fail' else '' end
    ,[701 Result] = case when gd.WashResult = 'P' then 'Pass' when gd.WashResult = 'F' then 'Fail' else '' end
    ,gd.inspector
    ,[Comments] = gd.Remark
    ,gdi.TestBeforePicture
    ,gdi.TestAfterPicture
from GarmentTest g WITH(NOLOCK)
inner join GarmentTest_Detail gd WITH(NOLOCK) on g.ID = gd.ID
left join [ExtendServer].PMSFile.dbo.GarmentTest_Detail gdi WITH(NOLOCK) on gd.ID=gdi.ID AND gd.No = gdi.No
outer apply(
	select Value =  r.Name 
	from Style s WITH(NOLOCK)
	inner join Reason r on s.SpecialMark = r.ID and r.ReasonTypeID = 'Style_SpecialMark'
	where s.ID = g.StyleID
	and s.BrandID = g.BrandID
	and s.SeasonID = g.SeasonID
)SpecialMark
where gd.ID = @ID and gd.No = @No
";
            return ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);
        }

        public bool Update_GarmentTestDetail_Result_Amend(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };

            string sqlcmd = @"
update gd
set Result = ''
,SeamBreakageResult = ''
,OdourResult = ''
,WashResult = ''
from GarmentTest_Detail gd  WITH(NOLOCK)
where gd.id=@ID and No=@No
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public bool Update_GarmentTestDetail_Result(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@_ID", DbType.String, ID } ,
                { "@_No", DbType.String, No } ,
            };

            string sqlcmd = @"
Declare @ID bigint = @_ID
Declare @No int = @_No

-- FGPT
select TestName,Result,ResultCnt = count(*)
into #tmpFGPTResult
from (
select s1.*,Result =  
	CASE WHEN  s1.TestUnit = 'N' AND s1.[TestResult] !='' THEN IIF( Cast( s1.[TestResult] as float) >= cast( s1.Criteria as float) ,'Pass' ,'Fail')
		 WHEN  s1.TestUnit = 'mm' THEN IIF(  s1.[TestResult] = '<=4' OR s1.[TestResult] = '≦4','Pass' , IIF( s1.[TestResult]='>4','Fail','')  )
		 WHEN  s1.TestUnit = 'Pass/Fail'  THEN s1.[TestResult]
		 ELSE '' 
	END 
	From GarmentTest_Detail_FGPT s1 WITH(NOLOCK)
	left join DropDownList ddl with (nolock) on  ddl.Type = 'PMS_FGPT_TestName' and ddl.ID = s1.TestName
	where s1.ID = @ID and No = @No
) a
group by TestName,Result
order by TestName

--FGWT
select Result,ResultCnt = count(*)
into #tmpFGWTResult
from (
select s1.*,Result =  
	IIF(Scale IS NOT NULL
    ,IIF(Scale='4-5' OR Scale ='5','Pass',IIF(Scale='','','Fail'))
    ,IIF( (BeforeWash IS NOT NULL AND AfterWash IS NOT NULL AND Criteria IS NOT NULL AND Shrinkage IS NOT NULL)
          or (Type = 'spirality: Garment - in percentage (average)')
          or (Type = 'spirality: Garment - in percentage (average) (Top Method A)')
          or (Type = 'spirality: Garment - in percentage (average) (Top Method B)')
          or (Type = 'spirality: Garment - in percentage (average) (Bottom Method A)')
          or (Type = 'spirality: Garment - in percentage (average) (Bottom Method B)')
   ,( IIF( TestDetail = '%' OR TestDetail = 'Range%'   
   -- % 為ISP20201331舊資料、Range% 為ISP20201606加上的新資料，兩者都視作百分比
      ---- 百分比 判斷方式
      ,IIF( ISNULL(Criteria,0)  <= ISNULL(Shrinkage,0) AND ISNULL(Shrinkage,0) <= ISNULL(Criteria2,0)
       , 'Pass'
       , 'Fail'
      )
      ---- 非百分比 判斷方式
      ,IIF( ISNULL(AfterWash,0) - ISNULL(BeforeWash,0) <= ISNULL(Criteria,0)
       ,'Pass'
       ,'Fail'
      )
    )
   )
   ,''
 )
)
	From GarmentTest_Detail_FGWT s1 WITH(NOLOCK)
	where s1.ID = @ID and No = @No
) a
group by Result

update gd
set
 SeamBreakageResult = case
	when NonSeamBreakageTest = 1 then ''
	when NonSeamBreakageTest = 0 and (select count(1) from #tmpFGPTResult where TestName = 'PHX-AP0450' and Result = 'Fail') > 0 then 'F'
	when NonSeamBreakageTest = 0 and (select count(1) from #tmpFGPTResult where TestName = 'PHX-AP0450' and Result = 'Pass') > 0 then 'P'
	when NonSeamBreakageTest = 0 and (select count(1) from #tmpFGPTResult where TestName = 'PHX-AP0450' and Result = '') > 0 then ''
	else ''  End
,OdourResult = case
	when  (select count(1) from #tmpFGPTResult where TestName = 'PHX-AP0451' and Result = 'Fail') > 0 then 'F'
	when  (select count(1) from #tmpFGPTResult where TestName = 'PHX-AP0451' and Result = 'Pass') > 0 then 'P'
	when  (select count(1) from #tmpFGPTResult where TestName = 'PHX-AP0451' and Result = '') > 0 then ''
	else ''  End
,WashResult = case
	when  (select count(1) from #tmpFGWTResult where Result = 'Fail') > 0 then 'F'
	when  (select count(1) from #tmpFGWTResult where Result = 'Pass') > 0 then 'P'
	when  (select count(1) from #tmpFGWTResult where Result = '') > 0 then ''
	else ''  End
from GarmentTest_Detail gd  WITH(NOLOCK)
where gd.id=@ID and No=@No

update gd
set
 Result = case 
	when gd.SeamBreakageResult = 'F' or gd.OdourResult = 'F' or gd.WashResult = 'F' then 'F'	
	when OdourResult = 'P' and WashResult = 'P' and (NonSeamBreakageTest = 1 or (NonSeamBreakageTest = 0 and SeamBreakageResult = 'P')) then 'P'
	when OdourResult = '' or WashResult = '' then '' 
	when (NonSeamBreakageTest = 0 and SeamBreakageResult = '') then ''
	else '' end
from GarmentTest_Detail gd  WITH(NOLOCK)
where gd.id=@ID and No=@No

drop table #tmpFGPTResult,#tmpFGWTResult

";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        /*回傳Garment Test(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Garment Test
        /// </summary>
        /// <param name="Item">Garment Test成員</param>
        /// <returns>回傳Garment Test</returns>
        /// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public IList<GarmentTest_Detail_ViewModel> Get(string ID, string No)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         gd.ID" + Environment.NewLine);
            SbSql.Append("        ,gd.No" + Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,inspdate"+ Environment.NewLine);
            SbSql.Append("        ,inspector"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,Sender"+ Environment.NewLine);
            SbSql.Append("        ,SendDate"+ Environment.NewLine);
            SbSql.Append("        ,Receiver"+ Environment.NewLine);
            SbSql.Append("        ,ReceiveDate"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,OldUkey"+ Environment.NewLine);
            SbSql.Append("        ,SubmitDate"+ Environment.NewLine);
            SbSql.Append("        ,ArrivedQty"+ Environment.NewLine);
            SbSql.Append("        ,LineDry"+ Environment.NewLine);
            SbSql.Append("        ,Temperature"+ Environment.NewLine);
            SbSql.Append("        ,TumbleDry"+ Environment.NewLine);
            SbSql.Append("        ,Machine"+ Environment.NewLine);
            SbSql.Append("        ,HandWash"+ Environment.NewLine);
            SbSql.Append("        ,Composition"+ Environment.NewLine);
            SbSql.Append("        ,Neck"+ Environment.NewLine);
            SbSql.Append("        ,Status"+ Environment.NewLine);
            SbSql.Append("        ,SizeCode"+ Environment.NewLine);
            SbSql.Append("        ,LOtoFactory"+ Environment.NewLine);
            SbSql.Append("        ,MtlTypeID"+ Environment.NewLine);
            SbSql.Append("        ,Above50NaturalFibres"+ Environment.NewLine);
            SbSql.Append("        ,Above50SyntheticFibres"+ Environment.NewLine);
            SbSql.Append("        ,OrderID"+ Environment.NewLine);
            SbSql.Append("        ,NonSeamBreakageTest"+ Environment.NewLine);
            SbSql.Append("        ,SeamBreakageResult"+ Environment.NewLine);
            SbSql.Append("        ,OdourResult"+ Environment.NewLine);
            SbSql.Append("        ,WashResult"+ Environment.NewLine);
            SbSql.Append("        ,gdi.TestBeforePicture" + Environment.NewLine);
            SbSql.Append("        ,gdi.TestAfterPicture" + Environment.NewLine);
            SbSql.Append($@"FROM [GarmentTest_Detail] gd WITH(NOLOCK)
left join [ExtendServer].PMSFile.dbo.GarmentTest_Detail gdi WITH(NOLOCK) on gd.ID=gdi.ID AND gd.No = gdi.No

" + Environment.NewLine);
            SbSql.Append("where gd.ID = @ID" + Environment.NewLine);
            SbSql.Append("and gd.No = @No" + Environment.NewLine);

            return ExecuteList<GarmentTest_Detail_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        
	#endregion
    }
}
