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
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", source.ID } ,
                { "@No", source.No } ,
                { "@SubmitDate", DbType.Date, source.SubmitDate},
                { "@ArrivedQty", source.ArrivedQty } ,
                { "@LOtoFactory", source.LOtoFactory } ,
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
    EditDate = GetDate()
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
,[TestDate] = gd.inspdate
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
		/*建立Garment Test(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Garment Test
        /// </summary>
        /// <param name="Item">Garment Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Create(GarmentTest_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [GarmentTest_Detail]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
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
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@No"); objParameter.Add("@No", DbType.Int32, Item.No);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@inspdate"); objParameter.Add("@inspdate", DbType.String, Item.inspdate);
            SbSql.Append("        ,@inspector"); objParameter.Add("@inspector", DbType.String, Item.inspector);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@Sender"); objParameter.Add("@Sender", DbType.String, Item.Sender);
            SbSql.Append("        ,@SendDate"); objParameter.Add("@SendDate", DbType.DateTime, Item.SendDate);
            SbSql.Append("        ,@Receiver"); objParameter.Add("@Receiver", DbType.String, Item.Receiver);
            SbSql.Append("        ,@ReceiveDate"); objParameter.Add("@ReceiveDate", DbType.DateTime, Item.ReceiveDate);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@OldUkey"); objParameter.Add("@OldUkey", DbType.String, Item.OldUkey);
            SbSql.Append("        ,@SubmitDate"); objParameter.Add("@SubmitDate", DbType.String, Item.SubmitDate);
            SbSql.Append("        ,@ArrivedQty"); objParameter.Add("@ArrivedQty", DbType.Int32, Item.ArrivedQty);
            SbSql.Append("        ,@LineDry"); objParameter.Add("@LineDry", DbType.String, Item.LineDry);
            SbSql.Append("        ,@Temperature"); objParameter.Add("@Temperature", DbType.Int32, Item.Temperature);
            SbSql.Append("        ,@TumbleDry"); objParameter.Add("@TumbleDry", DbType.String, Item.TumbleDry);
            SbSql.Append("        ,@Machine"); objParameter.Add("@Machine", DbType.String, Item.Machine);
            SbSql.Append("        ,@HandWash"); objParameter.Add("@HandWash", DbType.String, Item.HandWash);
            SbSql.Append("        ,@Composition"); objParameter.Add("@Composition", DbType.String, Item.Composition);
            SbSql.Append("        ,@Neck"); objParameter.Add("@Neck", DbType.String, Item.Neck);
            SbSql.Append("        ,@Status"); objParameter.Add("@Status", DbType.String, Item.Status);
            SbSql.Append("        ,@SizeCode"); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);
            SbSql.Append("        ,@LOtoFactory"); objParameter.Add("@LOtoFactory", DbType.String, Item.LOtoFactory);
            SbSql.Append("        ,@MtlTypeID"); objParameter.Add("@MtlTypeID", DbType.String, Item.MtlTypeID);
            SbSql.Append("        ,@Above50NaturalFibres"); objParameter.Add("@Above50NaturalFibres", DbType.String, Item.Above50NaturalFibres);
            SbSql.Append("        ,@Above50SyntheticFibres"); objParameter.Add("@Above50SyntheticFibres", DbType.String, Item.Above50SyntheticFibres);
            SbSql.Append("        ,@OrderID"); objParameter.Add("@OrderID", DbType.String, Item.OrderID);
            SbSql.Append("        ,@NonSeamBreakageTest"); objParameter.Add("@NonSeamBreakageTest", DbType.String, Item.NonSeamBreakageTest);
            SbSql.Append("        ,@SeamBreakageResult"); objParameter.Add("@SeamBreakageResult", DbType.String, Item.SeamBreakageResult);
            SbSql.Append("        ,@OdourResult"); objParameter.Add("@OdourResult", DbType.String, Item.OdourResult);
            SbSql.Append("        ,@WashResult"); objParameter.Add("@WashResult", DbType.String, Item.WashResult);
            SbSql.Append("        ,@TestBeforePicture"); objParameter.Add("@TestBeforePicture", DbType.String, Item.TestBeforePicture);
            SbSql.Append("        ,@TestAfterPicture"); objParameter.Add("@TestAfterPicture", DbType.String, Item.TestAfterPicture);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Garment Test(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Garment Test
        /// </summary>
        /// <param name="Item">Garment Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Update(GarmentTest_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [GarmentTest_Detail]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.No != null) { SbSql.Append(",No=@No"+ Environment.NewLine); objParameter.Add("@No", DbType.Int32, Item.No);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.inspdate != null) { SbSql.Append(",inspdate=@inspdate"+ Environment.NewLine); objParameter.Add("@inspdate", DbType.String, Item.inspdate);}
            if (Item.inspector != null) { SbSql.Append(",inspector=@inspector"+ Environment.NewLine); objParameter.Add("@inspector", DbType.String, Item.inspector);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.Sender != null) { SbSql.Append(",Sender=@Sender"+ Environment.NewLine); objParameter.Add("@Sender", DbType.String, Item.Sender);}
            if (Item.SendDate != null) { SbSql.Append(",SendDate=@SendDate"+ Environment.NewLine); objParameter.Add("@SendDate", DbType.DateTime, Item.SendDate);}
            if (Item.Receiver != null) { SbSql.Append(",Receiver=@Receiver"+ Environment.NewLine); objParameter.Add("@Receiver", DbType.String, Item.Receiver);}
            if (Item.ReceiveDate != null) { SbSql.Append(",ReceiveDate=@ReceiveDate"+ Environment.NewLine); objParameter.Add("@ReceiveDate", DbType.DateTime, Item.ReceiveDate);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.OldUkey != null) { SbSql.Append(",OldUkey=@OldUkey"+ Environment.NewLine); objParameter.Add("@OldUkey", DbType.String, Item.OldUkey);}
            if (Item.SubmitDate != null) { SbSql.Append(",SubmitDate=@SubmitDate"+ Environment.NewLine); objParameter.Add("@SubmitDate", DbType.String, Item.SubmitDate);}
            if (Item.ArrivedQty != null) { SbSql.Append(",ArrivedQty=@ArrivedQty"+ Environment.NewLine); objParameter.Add("@ArrivedQty", DbType.Int32, Item.ArrivedQty);}
            if (Item.LineDry != null) { SbSql.Append(",LineDry=@LineDry"+ Environment.NewLine); objParameter.Add("@LineDry", DbType.String, Item.LineDry);}
            if (Item.Temperature != null) { SbSql.Append(",Temperature=@Temperature"+ Environment.NewLine); objParameter.Add("@Temperature", DbType.Int32, Item.Temperature);}
            if (Item.TumbleDry != null) { SbSql.Append(",TumbleDry=@TumbleDry"+ Environment.NewLine); objParameter.Add("@TumbleDry", DbType.String, Item.TumbleDry);}
            if (Item.Machine != null) { SbSql.Append(",Machine=@Machine"+ Environment.NewLine); objParameter.Add("@Machine", DbType.String, Item.Machine);}
            if (Item.HandWash != null) { SbSql.Append(",HandWash=@HandWash"+ Environment.NewLine); objParameter.Add("@HandWash", DbType.String, Item.HandWash);}
            if (Item.Composition != null) { SbSql.Append(",Composition=@Composition"+ Environment.NewLine); objParameter.Add("@Composition", DbType.String, Item.Composition);}
            if (Item.Neck != null) { SbSql.Append(",Neck=@Neck"+ Environment.NewLine); objParameter.Add("@Neck", DbType.String, Item.Neck);}
            if (Item.Status != null) { SbSql.Append(",Status=@Status"+ Environment.NewLine); objParameter.Add("@Status", DbType.String, Item.Status);}
            if (Item.SizeCode != null) { SbSql.Append(",SizeCode=@SizeCode"+ Environment.NewLine); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);}
            if (Item.LOtoFactory != null) { SbSql.Append(",LOtoFactory=@LOtoFactory"+ Environment.NewLine); objParameter.Add("@LOtoFactory", DbType.String, Item.LOtoFactory);}
            if (Item.MtlTypeID != null) { SbSql.Append(",MtlTypeID=@MtlTypeID"+ Environment.NewLine); objParameter.Add("@MtlTypeID", DbType.String, Item.MtlTypeID);}
            if (Item.Above50NaturalFibres != null) { SbSql.Append(",Above50NaturalFibres=@Above50NaturalFibres"+ Environment.NewLine); objParameter.Add("@Above50NaturalFibres", DbType.String, Item.Above50NaturalFibres);}
            if (Item.Above50SyntheticFibres != null) { SbSql.Append(",Above50SyntheticFibres=@Above50SyntheticFibres"+ Environment.NewLine); objParameter.Add("@Above50SyntheticFibres", DbType.String, Item.Above50SyntheticFibres);}
            if (Item.OrderID != null) { SbSql.Append(",OrderID=@OrderID"+ Environment.NewLine); objParameter.Add("@OrderID", DbType.String, Item.OrderID);}
            if (Item.NonSeamBreakageTest != null) { SbSql.Append(",NonSeamBreakageTest=@NonSeamBreakageTest"+ Environment.NewLine); objParameter.Add("@NonSeamBreakageTest", DbType.String, Item.NonSeamBreakageTest);}
            if (Item.SeamBreakageResult != null) { SbSql.Append(",SeamBreakageResult=@SeamBreakageResult"+ Environment.NewLine); objParameter.Add("@SeamBreakageResult", DbType.String, Item.SeamBreakageResult);}
            if (Item.OdourResult != null) { SbSql.Append(",OdourResult=@OdourResult"+ Environment.NewLine); objParameter.Add("@OdourResult", DbType.String, Item.OdourResult);}
            if (Item.WashResult != null) { SbSql.Append(",WashResult=@WashResult"+ Environment.NewLine); objParameter.Add("@WashResult", DbType.String, Item.WashResult);}
            if (Item.TestBeforePicture != null) { SbSql.Append(",TestBeforePicture=@TestBeforePicture"+ Environment.NewLine); objParameter.Add("@TestBeforePicture", DbType.String, Item.TestBeforePicture);}
            if (Item.TestAfterPicture != null) { SbSql.Append(",TestAfterPicture=@TestAfterPicture"+ Environment.NewLine); objParameter.Add("@TestAfterPicture", DbType.String, Item.TestAfterPicture);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Garment Test(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Garment Test
        /// </summary>
        /// <param name="Item">Garment Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Delete(GarmentTest_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [GarmentTest_Detail]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
