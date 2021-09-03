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
                sqlcmd = "select distinct SizeCode from Style_SizeCode order by SizeCode";
            }
            else
            {
                sqlcmd = @"
select distinct SizeCode from Order_Qty
where ID = @OrderID and Article = @Article 
order by SizeCode";
            }

            return ExecuteList<Order_Qty>(CommandType.Text, sqlcmd, objParameter);
        }

        public List<string> GetScales()
        {
            string sqlcmd = @"select ID from Scale  WHERE Junk=0 order by ID";
            DataTable dt = ExecuteDataTable(CommandType.Text, sqlcmd, new SQLParameterCollection());

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
,gd.TestAfterPicture,gd.TestBeforePicture
from GarmentTest_Detail gd
left join Pass1 CreatBy on CreatBy.ID = gd.AddName
left join Pass1 EditBy on EditBy.ID = gd.EditName
outer apply(
	select Name_Extno 
	from View_ShowName
	where ID=gd.inspector
)InspectorName
where gd.ID = @ID
" + Environment.NewLine;
            

            return ExecuteList<GarmentTest_Detail_ViewModel>(CommandType.Text, sqlcmd, objParameter);
        }

        public int Delete_GarmentTestDetail(string ID, string No)
        {  
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };

            string sqlcmd = @"
Delete GarmentTest_Detail_Shrinkage  where id = @ID and NO = @No
Delete GarmentTest_Detail_Apperance where id = @ID and NO = @No
Delete GarmentTest_Detail_FGWT where id = @ID and NO = @No
Delete GarmentTest_Detail_FGPT where id = @ID and NO = @No
Delete Garment_Detail_Spirality where id = @ID and NO = @No

Delete GarmentTest_Detail where id = @ID and NO = @No
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
from GarmentTest_Detail_FGPT t
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
from GarmentTest_Detail_FGWT f
where ID = @ID
and No = @No

--
";
            return ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter).AsEnumerable().Where(x => x.Equals("Fail")).Any();
        }

        public string Get_LastResult(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            string sqlcmd = @"
select Result 
from  GarmentTest_Detail t
where No = (
	select Max(No) from GarmentTest_Detail s
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
                { "@Result", source.Result } ,
                { "@Remark", source.Remark } ,
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
            };

            if (source.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", source.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }

            if (source.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", source.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            string sqlcmd = $@"
update GarmentTest_Detail set
    SubmitDate = @SubmitDate,
    ArrivedQty =  @ArrivedQty,
    LOtoFactory =  @LOtoFactory,
    Result = @Result,
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
    EditName = @EditName,
    EditDate = GetDate(),
    TestBeforePicture = @TestBeforePicture,
    TestAfterPicture = @TestAfterPicture
where ID = @ID and No = @No
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
update GarmentTest_Detail 
set
    TestBeforePicture = @TestBeforePicture,
    TestAfterPicture = @TestAfterPicture
where ID = @ID and No = @No
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public bool Encode_GarmentTestDetail(string ID, string Status)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@Status", DbType.String, Status } ,
            };

            string sqlcmd = @"
Update GarmentTest_Detail set Status=@Status where id = @ID
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
,g.Result
,[450 Result] = gd.SeamBreakageResult
,[451 Result] = gd.OdourResult
,[701 Result] = gd.WashResult
,gd.inspector
,[Comments] = gd.Remark
from GarmentTest g
inner join GarmentTest_Detail gd on g.ID = gd.ID
outer apply(
	select Value =  r.Name 
	from Style s
	inner join Reason r on s.SpecialMark = r.ID and r.ReasonTypeID = 'Style_SpecialMark'
	where s.ID = g.StyleID
	and s.BrandID = g.BrandID
	and s.SeasonID = g.SeasonID
)SpecialMark
where gd.ID = @ID and gd.No = @No
";
            return ExecuteDataTable(CommandType.Text, sqlcmd, objParameter);
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
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
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
            SbSql.Append("        ,TestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture"+ Environment.NewLine);
            SbSql.Append("FROM [GarmentTest_Detail]"+ Environment.NewLine);
            SbSql.Append("where ID = @ID" + Environment.NewLine);
            SbSql.Append("and No = @No" + Environment.NewLine);

            return ExecuteList<GarmentTest_Detail_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        
	#endregion
    }
}
