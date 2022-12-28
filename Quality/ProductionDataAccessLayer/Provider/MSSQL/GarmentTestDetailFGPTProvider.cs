using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using System.Data.SqlClient;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentTestDetailFGPTProvider : SQLDAL, IGarmentTestDetailFGPTProvider
    {
        #region 底層連線
        public GarmentTestDetailFGPTProvider(string ConString) : base(ConString) { }
        public GarmentTestDetailFGPTProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public bool Update_FGPT(List<GarmentTest_Detail_FGPT_ViewModel> source)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            int idx = 0;
            string sqlcmd = string.Empty;


            foreach (var item in source)
            {
                // Key
                objParameter.Add(new SqlParameter($"@ID{idx}", item.ID));
                objParameter.Add(new SqlParameter($"@No{idx}", item.No));
                objParameter.Add(new SqlParameter($"@Location{idx}", item.Location));
                objParameter.Add(new SqlParameter($"@Type{idx}", item.TypeOri));
                objParameter.Add(new SqlParameter($"@Seq{idx}", item.Seq));
                objParameter.Add(new SqlParameter($"@TestName{idx}", item.TestName));

                objParameter.Add(new SqlParameter($"@TestResult{idx}", item.TestResult));
                objParameter.Add(new SqlParameter($"@TestUnit{idx}", item.TestUnit));
                objParameter.Add(new SqlParameter($"@TypeSelection_Seq{idx}", item.TypeSelection_Seq));
                objParameter.Add(new SqlParameter($"@TypeSelection_VersionID{idx}", item.TypeSelection_VersionID));


                sqlcmd += $@"
update g 
set
    [TestResult]    = isnull(@TestResult{idx}, ''),
    [TestUnit]	    = @TestUnit{idx},
    [TypeSelection_Seq]	= @TypeSelection_Seq{idx},
    [TypeSelection_VersionID]	= @TypeSelection_VersionID{idx}
from GarmentTest_Detail_FGPT g WITH(NOLOCK)
left join DropDownList d WITH(NOLOCK) on d.ID = g.TestName and d.Type = 'PMS_FGPT_TestName' 
where g.ID = @ID{idx} and g.No = @No{idx} and g.Type = @Type{idx} and g.Location = @Location{idx} 
and g.Seq = @Seq{idx} 
and isnull(d.Description, g.TestName) = @TestName{idx} 
" + Environment.NewLine;

                idx++;
            }

            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public IList<GarmentTest_Detail_FGPT_ViewModel> Get_GarmentTest_Detail_FGPT(string ID, string No)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@No", DbType.String, No } ,
            };
            string sqlcmd = @"
select 
[ID]
,[No]
,[Location]
,[LocationText]= CASE WHEN Location='B' THEN 'Bottom'
						WHEN Location='T' THEN 'Top'
						WHEN Location='S' THEN 'Top+Bottom'
						ELSE Location
					END
,[Type] = IIF(t.TypeSelection_VersionID > 0, Replace(t.type, '{0}', ts.Code), t.type)
,[TypeOri] = t.type
,[TestName] = ISNULL(PMS_FGPT_TestName.Description, t.TestName)
,[TestDetail]
,[Criteria]
,[TestResult]
,StandardRemark
,[TestUnit]
,t.[Seq]
,[TypeSelection_VersionID]
,[TypeSelection_Seq]
,[Result] = CASE WHEN  t.TestUnit = 'N' AND t.[TestResult] !='' THEN IIF( Cast( t.[TestResult] as float) >= cast( t.Criteria as float) ,'Pass' ,'Fail')
  WHEN  t.TestUnit = 'mm' THEN IIF(  t.[TestResult] = '<=4' OR t.[TestResult] = '≦4','Pass' , IIF( t.[TestResult]='>4','Fail','')  )
  WHEN  t.TestUnit = 'Pass/Fail'  THEN t.[TestResult]
   ELSE ''
END
,t.IsOriginal
from GarmentTest_Detail_FGPT t WITH(NOLOCK)
left join TypeSelection ts WITH(NOLOCK) on ts.VersionID = t.TypeSelection_VersionID 
        and ts.Seq = t.TypeSelection_Seq
outer apply(
	select Description
	from DropDownList  WITH(NOLOCK)
	where Type = 'PMS_FGPT_TestName' 
	and ID = t.TestName
)PMS_FGPT_TestName
where t.ID = @ID
and t.No = @No
order by PMS_FGPT_TestName.Description,t.[Seq]
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
            SbSql.Append("FROM [GarmentTest_Detail_FGPT] WITH(NOLOCK)" + Environment.NewLine);



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
