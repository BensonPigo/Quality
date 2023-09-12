using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using Newtonsoft.Json;
using DatabaseObject;
using DatabaseObject.ViewModel.FinalInspection;
using System.Transactions;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class FinalInspectionMeasurementProvider : SQLDAL, IFinalInspection_MeasurementProvider
    {
        #region 底層連線
        public FinalInspectionMeasurementProvider(string conString) : base(conString) { }
        public FinalInspectionMeasurementProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<QueryReport_Measurement> GetQuery_FinalInspection_Measurement(string ID, string SP)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection() {
            { "@ID", DbType.String, ID },
            { "@SP", DbType.String, SP }
            };

            SbSql.Append($@"
SELECT
	Time = format(fm.AddDate,'HH:mm:ss'),
	fm.[Article],
	fm.[SizeCode],
	fm.[Location],
	m.Description,
	Tol2 = iif(m.Tol2 = '', '0', m.Tol2),
	Tol1 = iif(m.Tol1 = '', '0', m.Tol1),
	m.SizeSpec,
	SizeSpec2 = fm.[SizeSpec],
	[diff] = dbo.calculateSizeSpec(m.SizeSpec, fm.SizeSpec, SizeUnit.SizeUnit)
FROM [FinalInspection_Measurement] fm WITH(NOLOCK)
inner join FinalInspection f WITH(NOLOCK) on f.ID = fm.ID
inner join Measurement m WITH(NOLOCK) on m.Ukey = fm.MeasurementUkey
inner join Production.dbo.Style s with(nolock) on s.ukey = m.StyleUkey
left join Production.dbo.Orders o1 with(nolock) on o1.id = @SP
outer apply(
	select top 1 SizeUnit from Production.dbo.Orders o WITH(NOLOCK)
	where StyleUkey = m.StyleUkey and id = @SP
		and LocalOrder=1 and SubconInSisterFty=1
)o2
outer apply(select SizeUnit = iif(len(o1.SizeUnit) > 0,o1.SizeUnit , o2.SizeUnit)) SizeUnit
where fm.ID = @ID
AND (m.SizeSpec NOT LIKE '%!%' AND m.SizeSpec NOT LIKE '%@%' AND m.SizeSpec NOT LIKE '%#%' 
AND m.SizeSpec NOT LIKE '%$%'  AND m.SizeSpec NOT LIKE '%^%'  AND m.SizeSpec NOT LIKE '%&%' 
AND m.SizeSpec NOT LIKE '%*%' AND m.SizeSpec NOT LIKE '%=%' AND m.SizeSpec NOT LIKE '%-%' 
AND m.SizeSpec NOT LIKE '%(%' AND m.SizeSpec NOT LIKE '%)%')

order by Time,fm.[Article],fm.[SizeCode],fm.Location
");
            return ExecuteList<QueryReport_Measurement>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        #region CRUD Base
        /*回傳(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳
        /// </summary>
        /// <param name="Item">成員</param>
        /// <returns>回傳</returns>
        /// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public IList<FinalInspection_Measurement> Get(FinalInspection_Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         Ukey" + Environment.NewLine);
            SbSql.Append("        ,ID" + Environment.NewLine);
            SbSql.Append("        ,Article" + Environment.NewLine);
            SbSql.Append("        ,SizeCode" + Environment.NewLine);
            SbSql.Append("        ,Location" + Environment.NewLine);
            SbSql.Append("        ,Code" + Environment.NewLine);
            SbSql.Append("        ,SizeSpec" + Environment.NewLine);
            SbSql.Append("        ,MeasurementUkey" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("FROM [FinalInspection_Measurement] WITH(NOLOCK)" + Environment.NewLine);



            return ExecuteList<FinalInspection_Measurement>(CommandType.Text, SbSql.ToString(), objParameter);
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
        public int Create(FinalInspection_Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [FinalInspection_Measurement]" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         Ukey" + Environment.NewLine);
            SbSql.Append("        ,ID" + Environment.NewLine);
            SbSql.Append("        ,Article" + Environment.NewLine);
            SbSql.Append("        ,SizeCode" + Environment.NewLine);
            SbSql.Append("        ,Location" + Environment.NewLine);
            SbSql.Append("        ,Code" + Environment.NewLine);
            SbSql.Append("        ,SizeSpec" + Environment.NewLine);
            SbSql.Append("        ,MeasurementUkey" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append(")" + Environment.NewLine);
            SbSql.Append("VALUES" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @Ukey"); objParameter.Add("@Ukey", DbType.String, Item.Ukey);
            SbSql.Append("        ,@ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@SizeCode"); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);
            SbSql.Append("        ,@Location"); objParameter.Add("@Location", DbType.String, Item.Location);
            SbSql.Append("        ,@Code"); objParameter.Add("@Code", DbType.String, Item.Code);
            SbSql.Append("        ,@SizeSpec"); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);
            SbSql.Append("        ,@MeasurementUkey"); objParameter.Add("@MeasurementUkey", DbType.String, Item.MeasurementUkey);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append(")" + Environment.NewLine);




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
        public int Update(FinalInspection_Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [FinalInspection_Measurement]" + Environment.NewLine);
            SbSql.Append("SET" + Environment.NewLine);
            if (Item.Ukey != null) { SbSql.Append("Ukey=@Ukey" + Environment.NewLine); objParameter.Add("@Ukey", DbType.String, Item.Ukey); }
            if (Item.ID != null) { SbSql.Append(",ID=@ID" + Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID); }
            if (Item.Article != null) { SbSql.Append(",Article=@Article" + Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article); }
            if (Item.SizeCode != null) { SbSql.Append(",SizeCode=@SizeCode" + Environment.NewLine); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode); }
            if (Item.Location != null) { SbSql.Append(",Location=@Location" + Environment.NewLine); objParameter.Add("@Location", DbType.String, Item.Location); }
            if (Item.Code != null) { SbSql.Append(",Code=@Code" + Environment.NewLine); objParameter.Add("@Code", DbType.String, Item.Code); }
            if (Item.SizeSpec != null) { SbSql.Append(",SizeSpec=@SizeSpec" + Environment.NewLine); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec); }
            if (Item.MeasurementUkey != null) { SbSql.Append(",MeasurementUkey=@MeasurementUkey" + Environment.NewLine); objParameter.Add("@MeasurementUkey", DbType.String, Item.MeasurementUkey); }
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName" + Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName); }
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
        /// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Delete(FinalInspection_Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [FinalInspection_Measurement]" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion
    }
}
