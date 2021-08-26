using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentTestProvider : SQLDAL, IGarmentTestProvider
    {
        #region 底層連線
        public GarmentTestProvider(string ConString) : base(ConString) { }
        public GarmentTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public IList<Style> GetStyleID()
        {
            string sqlcmd = @"select distinct ID from Style where Junk = 0";

            return ExecuteList<Style>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }

        public IList<Brand> GetBrandID()
        {
            string sqlcmd = @"select distinct ID from Brand where Junk = 0";

            return ExecuteList<Brand>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }

        public IList<Season> GetSeasonID()
        {
            string sqlcmd = @"select distinct ID from Season where Junk = 0";

            return ExecuteList<Season>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }

        public IList<GarmentTest> GetArticle(GarmentTest_ViewModel filter)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, filter.BrandID } ,
                { "@StyleID", DbType.String, filter.StyleID } ,
                { "@SeasonID", DbType.String, filter.SeasonID} ,
            };
            string sqlcmd = @"
select distinct Article from GarmentTest
where 1=1
and BrandID = @BrandID
and StyleID = @StyleID
and SeasonID = @SeasonID";

            return ExecuteList<GarmentTest>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<GarmentTest_ViewModel> Get_GarmentTest(GarmentTest_ViewModel filter)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, filter.BrandID } ,
                { "@StyleID", DbType.String, filter.StyleID } ,
                { "@SeasonID", DbType.String, filter.SeasonID} ,
                { "@Article", DbType.String, filter.Article} ,
            };
            string sqlcmd = @"
select g.ID
,g.StyleID
,g.BrandID
,g.Article
,g.SeasonID
,g.FirstOrderID
,g.DeadLine
,g.SewingInline
,g.OrderID
,g.MDivisionid
,[MinSciDelivery] = GetSCI.MinSciDelivery
,[MinBuyerDelivery] = GetSCI.MinBuyerDelivery
,[SeamBreakageResult] = case when g.SeamBreakageResult = 'P' then 'Pass'
						     when g.SeamBreakageResult = 'F' then 'Fail' 
                             else '' end
,[OdourResult] = case when g.OdourResult = 'P' then 'Pass'
					  when g.OdourResult = 'F' then 'Fail' 
                      else '' end
,[WashResult] = case when g.WashResult = 'P' then 'Pass'
				     when g.WashResult = 'F' then 'Fail' 
                     else '' end
,[WashName] = IIF(WashName.Value is null,'701','710')
,[SpecialMark] = SpecialMark.Value
,g.Result, g.Date,g.Remark
,[GarmentTestAddName] = CONCAT(g.AddName,'-',CreatBy.Name,'',g.AddDate)
,[GarmentTestEditName] = CONCAT(g.EditName,'-',EditBy.Name,'',g.EditDate)
,g.AddName,g.EditName
from GarmentTest g
left join Pass1 CreatBy on CreatBy.ID = g.AddName
left join Pass1 EditBy on EditBy.ID = g.EditName
outer apply(
	select MinBuyerDelivery,MinSciDelivery
	from dbo.GetSCI(g.FirstOrderID,'')
) GetSCI
outer apply(
	select Value =  r.Name 
	from Style s
	inner join Reason r on s.SpecialMark = r.ID and r.ReasonTypeID = 'Style_SpecialMark'
	where s.ID = g.StyleID
	and s.BrandID = g.BrandID
	and s.SeasonID = g.SeasonID
)SpecialMark
outer apply(
	select Value =  r.Name 
	from Style s
	inner join Reason r on s.SpecialMark = r.ID and r.ReasonTypeID = 'Style_SpecialMark'
	where s.ID = g.StyleID
	and s.BrandID = g.BrandID
	and s.SeasonID = g.SeasonID
	and r.Name in ('MATCH TEAMWEAR','BASEBALL ON FIELD','SOFTBALL ON FIELD','TRAINING TEAMWEAR','LACROSSE ONFIELD','AMERIC. FOOT. ON-FIELD','TIRO','BASEBALL OFF FIELD','NCAA ON ICE','ON-COURT','BBALL PERFORMANCE','BRANDED BLANKS','SLD ON-FIELD','NHL ON ICE','SLD ON-COURT')
)WashName
where 1=1
and g.BrandID = @BrandID
and g.StyleID = @StyleID
and g.SeasonID = @SeasonID
and g.Article = @Article" + Environment.NewLine;
            if (!string.IsNullOrEmpty(filter.MDivisionid))
            {
                objParameter.Add("@MDivisionid", DbType.String, filter.MDivisionid);
                sqlcmd += " and MDivisionid = @MDivisionid";
            }

            return ExecuteList<GarmentTest_ViewModel>(CommandType.Text, sqlcmd, objParameter);
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
        public IList<GarmentTest> Get(GarmentTest Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,FirstOrderID"+ Environment.NewLine);
            SbSql.Append("        ,OrderID"+ Environment.NewLine);
            SbSql.Append("        ,StyleID"+ Environment.NewLine);
            SbSql.Append("        ,SeasonID"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,MDivisionid"+ Environment.NewLine);
            SbSql.Append("        ,DeadLine"+ Environment.NewLine);
            SbSql.Append("        ,SewingInline"+ Environment.NewLine);
            SbSql.Append("        ,SewingOffline"+ Environment.NewLine);
            SbSql.Append("        ,Date"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,OldUkey"+ Environment.NewLine);
            SbSql.Append("        ,SeamBreakageResult"+ Environment.NewLine);
            SbSql.Append("        ,SeamBreakageLastTestDate"+ Environment.NewLine);
            SbSql.Append("        ,OdourResult"+ Environment.NewLine);
            SbSql.Append("        ,WashResult"+ Environment.NewLine);
            SbSql.Append("FROM [GarmentTest]"+ Environment.NewLine);



            return ExecuteList<GarmentTest>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(GarmentTest Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [GarmentTest]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,FirstOrderID"+ Environment.NewLine);
            SbSql.Append("        ,OrderID"+ Environment.NewLine);
            SbSql.Append("        ,StyleID"+ Environment.NewLine);
            SbSql.Append("        ,SeasonID"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,MDivisionid"+ Environment.NewLine);
            SbSql.Append("        ,DeadLine"+ Environment.NewLine);
            SbSql.Append("        ,SewingInline"+ Environment.NewLine);
            SbSql.Append("        ,SewingOffline"+ Environment.NewLine);
            SbSql.Append("        ,Date"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,OldUkey"+ Environment.NewLine);
            SbSql.Append("        ,SeamBreakageResult"+ Environment.NewLine);
            SbSql.Append("        ,SeamBreakageLastTestDate"+ Environment.NewLine);
            SbSql.Append("        ,OdourResult"+ Environment.NewLine);
            SbSql.Append("        ,WashResult"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@FirstOrderID"); objParameter.Add("@FirstOrderID", DbType.String, Item.FirstOrderID);
            SbSql.Append("        ,@OrderID"); objParameter.Add("@OrderID", DbType.String, Item.OrderID);
            SbSql.Append("        ,@StyleID"); objParameter.Add("@StyleID", DbType.String, Item.StyleID);
            SbSql.Append("        ,@SeasonID"); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
            SbSql.Append("        ,@BrandID"); objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@MDivisionid"); objParameter.Add("@MDivisionid", DbType.String, Item.MDivisionid);
            SbSql.Append("        ,@DeadLine"); objParameter.Add("@DeadLine", DbType.String, Item.DeadLine);
            SbSql.Append("        ,@SewingInline"); objParameter.Add("@SewingInline", DbType.String, Item.SewingInline);
            SbSql.Append("        ,@SewingOffline"); objParameter.Add("@SewingOffline", DbType.String, Item.SewingOffline);
            SbSql.Append("        ,@Date"); objParameter.Add("@Date", DbType.String, Item.Date);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@OldUkey"); objParameter.Add("@OldUkey", DbType.String, Item.OldUkey);
            SbSql.Append("        ,@SeamBreakageResult"); objParameter.Add("@SeamBreakageResult", DbType.String, Item.SeamBreakageResult);
            SbSql.Append("        ,@SeamBreakageLastTestDate"); objParameter.Add("@SeamBreakageLastTestDate", DbType.String, Item.SeamBreakageLastTestDate);
            SbSql.Append("        ,@OdourResult"); objParameter.Add("@OdourResult", DbType.String, Item.OdourResult);
            SbSql.Append("        ,@WashResult"); objParameter.Add("@WashResult", DbType.String, Item.WashResult);
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
        public int Update(GarmentTest Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [GarmentTest]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.FirstOrderID != null) { SbSql.Append(",FirstOrderID=@FirstOrderID"+ Environment.NewLine); objParameter.Add("@FirstOrderID", DbType.String, Item.FirstOrderID);}
            if (Item.OrderID != null) { SbSql.Append(",OrderID=@OrderID"+ Environment.NewLine); objParameter.Add("@OrderID", DbType.String, Item.OrderID);}
            if (Item.StyleID != null) { SbSql.Append(",StyleID=@StyleID"+ Environment.NewLine); objParameter.Add("@StyleID", DbType.String, Item.StyleID);}
            if (Item.SeasonID != null) { SbSql.Append(",SeasonID=@SeasonID"+ Environment.NewLine); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);}
            if (Item.BrandID != null) { SbSql.Append(",BrandID=@BrandID"+ Environment.NewLine); objParameter.Add("@BrandID", DbType.String, Item.BrandID);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.MDivisionid != null) { SbSql.Append(",MDivisionid=@MDivisionid"+ Environment.NewLine); objParameter.Add("@MDivisionid", DbType.String, Item.MDivisionid);}
            if (Item.DeadLine != null) { SbSql.Append(",DeadLine=@DeadLine"+ Environment.NewLine); objParameter.Add("@DeadLine", DbType.String, Item.DeadLine);}
            if (Item.SewingInline != null) { SbSql.Append(",SewingInline=@SewingInline"+ Environment.NewLine); objParameter.Add("@SewingInline", DbType.String, Item.SewingInline);}
            if (Item.SewingOffline != null) { SbSql.Append(",SewingOffline=@SewingOffline"+ Environment.NewLine); objParameter.Add("@SewingOffline", DbType.String, Item.SewingOffline);}
            if (Item.Date != null) { SbSql.Append(",Date=@Date"+ Environment.NewLine); objParameter.Add("@Date", DbType.String, Item.Date);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.OldUkey != null) { SbSql.Append(",OldUkey=@OldUkey"+ Environment.NewLine); objParameter.Add("@OldUkey", DbType.String, Item.OldUkey);}
            if (Item.SeamBreakageResult != null) { SbSql.Append(",SeamBreakageResult=@SeamBreakageResult"+ Environment.NewLine); objParameter.Add("@SeamBreakageResult", DbType.String, Item.SeamBreakageResult);}
            if (Item.SeamBreakageLastTestDate != null) { SbSql.Append(",SeamBreakageLastTestDate=@SeamBreakageLastTestDate"+ Environment.NewLine); objParameter.Add("@SeamBreakageLastTestDate", DbType.String, Item.SeamBreakageLastTestDate);}
            if (Item.OdourResult != null) { SbSql.Append(",OdourResult=@OdourResult"+ Environment.NewLine); objParameter.Add("@OdourResult", DbType.String, Item.OdourResult);}
            if (Item.WashResult != null) { SbSql.Append(",WashResult=@WashResult"+ Environment.NewLine); objParameter.Add("@WashResult", DbType.String, Item.WashResult);}
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
        public int Delete(GarmentTest Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [GarmentTest]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
