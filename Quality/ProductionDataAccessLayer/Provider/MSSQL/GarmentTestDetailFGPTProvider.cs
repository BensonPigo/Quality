using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentTestDetailFGPTProvider : SQLDAL, IGarmentTestDetailFGPTProvider
    {
        #region 底層連線
        public GarmentTestDetailFGPTProvider(string ConString) : base(ConString) { }
        public GarmentTestDetailFGPTProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public IList<GarmentTest_Detail_FGPT_ViewModel> Get_GarmentTest_Detail_FGPT(Int64 ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.Int64, ID } ,
                { "@No", DbType.String, No } ,
            };
            string sqlcmd = @"
select 
[ID]
,[No]
,[Location]
,[Type] = IIF(t.TypeSelection_VersionID > 0, Replace(t.type, '{0}', ts.Code), t.type)
,[TestName] = PMS_FGPT_TestName.Description
,[TestDetail]
,[Criteria]
,[TestResult]
,[TestUnit]
,t.[Seq]
,[TypeSelection_VersionID]
,[TypeSelection_Seq]
,[Result] = CASE WHEN  t.TestUnit = 'N' AND t.[TestResult] !='' THEN IIF( Cast( t.[TestResult] as float) >= cast( t.Criteria as float) ,'Pass' ,'Fail')
  WHEN  t.TestUnit = 'mm' THEN IIF(  t.[TestResult] = '<=4' OR t.[TestResult] = '≦4','Pass' , IIF( t.[TestResult]='>4','Fail','')  )
  WHEN  t.TestUnit = 'Pass/Fail'  THEN t.[TestResult]
   ELSE ''
END
from GarmentTest_Detail_FGPT t
left join TypeSelection ts on ts.VersionID = t.TypeSelection_VersionID 
        and ts.Seq = t.TypeSelection_Seq
outer apply(
	select Description
	from DropDownList 
	where Type = 'PMS_FGPT_TestName' 
	and ID = t.TestName
)PMS_FGPT_TestName
where t.ID = @ID
and t.No = @No
";
            return ExecuteList<GarmentTest_Detail_FGPT_ViewModel>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<GarmentTest_Detail_FGPT> Get(GarmentTest_Detail_FGPT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,TestName"+ Environment.NewLine);
            SbSql.Append("        ,TestDetail"+ Environment.NewLine);
            SbSql.Append("        ,Criteria"+ Environment.NewLine);
            SbSql.Append("        ,TestResult"+ Environment.NewLine);
            SbSql.Append("        ,TestUnit"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("        ,TypeSelection_VersionID"+ Environment.NewLine);
            SbSql.Append("        ,TypeSelection_Seq"+ Environment.NewLine);
            SbSql.Append("FROM [GarmentTest_Detail_FGPT]"+ Environment.NewLine);



            return ExecuteList<GarmentTest_Detail_FGPT>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Create(GarmentTest_Detail_FGPT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [GarmentTest_Detail_FGPT]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,Type"+ Environment.NewLine);
            SbSql.Append("        ,TestName"+ Environment.NewLine);
            SbSql.Append("        ,TestDetail"+ Environment.NewLine);
            SbSql.Append("        ,Criteria"+ Environment.NewLine);
            SbSql.Append("        ,TestResult"+ Environment.NewLine);
            SbSql.Append("        ,TestUnit"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("        ,TypeSelection_VersionID"+ Environment.NewLine);
            SbSql.Append("        ,TypeSelection_Seq"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@No"); objParameter.Add("@No", DbType.Int32, Item.No);
            SbSql.Append("        ,@Location"); objParameter.Add("@Location", DbType.String, Item.Location);
            SbSql.Append("        ,@Type"); objParameter.Add("@Type", DbType.String, Item.Type);
            SbSql.Append("        ,@TestName"); objParameter.Add("@TestName", DbType.String, Item.TestName);
            SbSql.Append("        ,@TestDetail"); objParameter.Add("@TestDetail", DbType.String, Item.TestDetail);
            SbSql.Append("        ,@Criteria"); objParameter.Add("@Criteria", DbType.Int32, Item.Criteria);
            SbSql.Append("        ,@TestResult"); objParameter.Add("@TestResult", DbType.String, Item.TestResult);
            SbSql.Append("        ,@TestUnit"); objParameter.Add("@TestUnit", DbType.String, Item.TestUnit);
            SbSql.Append("        ,@Seq"); objParameter.Add("@Seq", DbType.Int32, Item.Seq);
            SbSql.Append("        ,@TypeSelection_VersionID"); objParameter.Add("@TypeSelection_VersionID", DbType.Int32, Item.TypeSelection_VersionID);
            SbSql.Append("        ,@TypeSelection_Seq"); objParameter.Add("@TypeSelection_Seq", DbType.Int32, Item.TypeSelection_Seq);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Update(GarmentTest_Detail_FGPT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [GarmentTest_Detail_FGPT]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.No != null) { SbSql.Append(",No=@No"+ Environment.NewLine); objParameter.Add("@No", DbType.Int32, Item.No);}
            if (Item.Location != null) { SbSql.Append(",Location=@Location"+ Environment.NewLine); objParameter.Add("@Location", DbType.String, Item.Location);}
            if (Item.Type != null) { SbSql.Append(",Type=@Type"+ Environment.NewLine); objParameter.Add("@Type", DbType.String, Item.Type);}
            if (Item.TestName != null) { SbSql.Append(",TestName=@TestName"+ Environment.NewLine); objParameter.Add("@TestName", DbType.String, Item.TestName);}
            if (Item.TestDetail != null) { SbSql.Append(",TestDetail=@TestDetail"+ Environment.NewLine); objParameter.Add("@TestDetail", DbType.String, Item.TestDetail);}
            if (Item.Criteria != null) { SbSql.Append(",Criteria=@Criteria"+ Environment.NewLine); objParameter.Add("@Criteria", DbType.Int32, Item.Criteria);}
            if (Item.TestResult != null) { SbSql.Append(",TestResult=@TestResult"+ Environment.NewLine); objParameter.Add("@TestResult", DbType.String, Item.TestResult);}
            if (Item.TestUnit != null) { SbSql.Append(",TestUnit=@TestUnit"+ Environment.NewLine); objParameter.Add("@TestUnit", DbType.String, Item.TestUnit);}
            if (Item.Seq != null) { SbSql.Append(",Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.Int32, Item.Seq);}
            if (Item.TypeSelection_VersionID != null) { SbSql.Append(",TypeSelection_VersionID=@TypeSelection_VersionID"+ Environment.NewLine); objParameter.Add("@TypeSelection_VersionID", DbType.Int32, Item.TypeSelection_VersionID);}
            if (Item.TypeSelection_Seq != null) { SbSql.Append(",TypeSelection_Seq=@TypeSelection_Seq"+ Environment.NewLine); objParameter.Add("@TypeSelection_Seq", DbType.Int32, Item.TypeSelection_Seq);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Delete(GarmentTest_Detail_FGPT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [GarmentTest_Detail_FGPT]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
