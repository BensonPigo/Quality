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
    public class RFTInspectionMeasurementProvider : SQLDAL, IRFTInspectionMeasurementProvider
    {
        #region 底層連線
        public RFTInspectionMeasurementProvider(string conString) : base(conString) { }
        public RFTInspectionMeasurementProvider(SQLDataTransaction tra) : base(tra) { }
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
        public IList<RFT_Inspection_Measurement_ViewModel> Get( Int64 StyleUkey, string SizeCode, string UserID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@StyleUkey", DbType.Int64, StyleUkey } ,
                { "@SizeCode", DbType.String, SizeCode } ,
                { "@UserID", DbType.String, UserID } ,
            };


            string sqlcmd = @"

exec CopyStyle_ToMeasurement @UserID,@StyleUkey;

SELECT [MeasurementUkey] = a.ukey
	,a.StyleUkey
	,a.Code
	,a.SizeCode 
	,a.SizeSpec
	,[Description] = a.Description	
	,a.Tol1
	,a.Tol2
    , [IsPatternMeas] = 
		 case when a.Description like '%pattern measn%'  then convert(bit, 1)
			  when  isnull(a.Tol1, '0') = '0' and isnull(a.Tol2, '0') = '0' and isnull(a.SizeSpec, '0') = '0' then convert(bit, 1)
			  when  isnull(a.SizeCode,'') = '' then convert(bit, 1)
			  when  a.SizeSpec like '%[a-zA-Z]%' then convert(bit, 1)
			  when  (UPPER(s.SizeUnit) = 'INCH' and  a.SizeSpec like '%.%') then convert(bit, 1)
			  when  (UPPER(s.SizeUnit) = 'CM' and  a.SizeSpec like '%/%') then convert(bit, 1)
			  else  convert(bit, 0) end
    ,[SizeUnit] = s.SizeUnit
	,OrderID = '', Article = '', Location = '',Line = '',FactoryID = ''
FROM [ManufacturingExecution].[dbo].[Measurement] a with(nolock)
LEFT JOIN [ManufacturingExecution].[dbo].[MeasurementTranslate] b ON  a.MeasurementTranslateUkey = b.UKey
LEFT JOIN Production.dbo.Style s on s.Ukey = a.StyleUkey
--LEFT JOIN [ManufacturingExecution].[dbo].Inspection_Measurement inspm on inspm.MeasurementUkey = a.Ukey
where a.junk=0 
and a.StyleUkey = @StyleUkey 
and a.SizeCode = @SizeCode

";

            return ExecuteList<RFT_Inspection_Measurement_ViewModel>(CommandType.Text, sqlcmd, objParameter, 60);
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
        public int Create(RFT_Inspection_Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [RFT_Inspection_Measurement]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         Ukey"+ Environment.NewLine);
            SbSql.Append("        ,MeasurementUkey"+ Environment.NewLine);
            SbSql.Append("        ,StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,No"+ Environment.NewLine);
            SbSql.Append("        ,Code"+ Environment.NewLine);
            SbSql.Append("        ,SizeCode"+ Environment.NewLine);
            SbSql.Append("        ,SizeSpec"+ Environment.NewLine);
            SbSql.Append("        ,OrderID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,Line"+ Environment.NewLine);
            SbSql.Append("        ,FactoryID"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @Ukey"); objParameter.Add("@Ukey", DbType.String, Item.Ukey);
            SbSql.Append("        ,@MeasurementUkey"); objParameter.Add("@MeasurementUkey", DbType.String, Item.MeasurementUkey);
            SbSql.Append("        ,@StyleUkey"); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);
            SbSql.Append("        ,@No"); objParameter.Add("@No", DbType.Int32, Item.No);
            SbSql.Append("        ,@Code"); objParameter.Add("@Code", DbType.String, Item.Code);
            SbSql.Append("        ,@SizeCode"); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);
            SbSql.Append("        ,@SizeSpec"); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);
            SbSql.Append("        ,@OrderID"); objParameter.Add("@OrderID", DbType.String, Item.OrderID);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@Location"); objParameter.Add("@Location", DbType.String, Item.Location);
            SbSql.Append("        ,@Line"); objParameter.Add("@Line", DbType.String, Item.Line);
            SbSql.Append("        ,@FactoryID"); objParameter.Add("@FactoryID", DbType.String, Item.FactoryID);
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
        public int Update(RFT_Inspection_Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [RFT_Inspection_Measurement]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.Ukey != null) { SbSql.Append("Ukey=@Ukey"+ Environment.NewLine); objParameter.Add("@Ukey", DbType.String, Item.Ukey);}
            if (Item.MeasurementUkey != null) { SbSql.Append(",MeasurementUkey=@MeasurementUkey"+ Environment.NewLine); objParameter.Add("@MeasurementUkey", DbType.String, Item.MeasurementUkey);}
            if (Item.StyleUkey != null) { SbSql.Append(",StyleUkey=@StyleUkey"+ Environment.NewLine); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);}
            if (Item.No != null) { SbSql.Append(",No=@No"+ Environment.NewLine); objParameter.Add("@No", DbType.Int32, Item.No);}
            if (Item.Code != null) { SbSql.Append(",Code=@Code"+ Environment.NewLine); objParameter.Add("@Code", DbType.String, Item.Code);}
            if (Item.SizeCode != null) { SbSql.Append(",SizeCode=@SizeCode"+ Environment.NewLine); objParameter.Add("@SizeCode", DbType.String, Item.SizeCode);}
            if (Item.SizeSpec != null) { SbSql.Append(",SizeSpec=@SizeSpec"+ Environment.NewLine); objParameter.Add("@SizeSpec", DbType.String, Item.SizeSpec);}
            if (Item.OrderID != null) { SbSql.Append(",OrderID=@OrderID"+ Environment.NewLine); objParameter.Add("@OrderID", DbType.String, Item.OrderID);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.Location != null) { SbSql.Append(",Location=@Location"+ Environment.NewLine); objParameter.Add("@Location", DbType.String, Item.Location);}
            if (Item.Line != null) { SbSql.Append(",Line=@Line"+ Environment.NewLine); objParameter.Add("@Line", DbType.String, Item.Line);}
            if (Item.FactoryID != null) { SbSql.Append(",FactoryID=@FactoryID"+ Environment.NewLine); objParameter.Add("@FactoryID", DbType.String, Item.FactoryID);}
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
        public int Delete(RFT_Inspection_Measurement Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [RFT_Inspection_Measurement]"+ Environment.NewLine);

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Save(List<RFT_Inspection_Measurement> Measurement)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string strNo = string.Empty;

            DataTable dt = ExecuteDataTable(CommandType.Text, $@"
select no = isnull(max(no),0)+1 from  ManufacturingExecution.dbo.RFT_Inspection_Measurement  WITH (NOLOCK) where styleukey = {Measurement[0].StyleUkey}", objParameter);
            if (dt != null || dt.Rows.Count > 0)
            {
                strNo = dt.Rows[0]["no"].ToString();
            }

            int rowSeq = 1;
            string sqlcmd = " declare @AddDate datetime = GetDate()";
            foreach (var item in Measurement)
            {
                if (!string.IsNullOrEmpty(item.SizeSpec))
                {
                    objParameter.Add($"@MeasurementUkey{rowSeq}", item.MeasurementUkey);
                    objParameter.Add($"@StyleUkey{rowSeq}", item.StyleUkey);
                    objParameter.Add($"@No", strNo);
                    objParameter.Add($"@Code{rowSeq}", item.Code);
                    objParameter.Add($"@SizeCode{rowSeq}", item.SizeCode);
                    objParameter.Add($"@SizeSpec{rowSeq}", item.SizeSpec);
                    objParameter.Add($"@OrderID{rowSeq}", item.OrderID);
                    objParameter.Add($"@Article{rowSeq}", item.Article);
                    objParameter.Add($"@Location{rowSeq}", item.Location);
                    objParameter.Add($"@Line{rowSeq}", item.Line);
                    objParameter.Add($"@FactoryID{rowSeq}", item.FactoryID);

                    sqlcmd += $@"
insert into RFT_Inspection_Measurement(MeasurementUkey,StyleUkey,No,Code,SizeCode,SizeSpec,OrderID,Article,Location,Line,FactoryID,AddDate)
values(@MeasurementUkey{rowSeq},@StyleUkey{rowSeq},@No,@Code{rowSeq},@SizeCode{rowSeq},@SizeSpec{rowSeq},@OrderID{rowSeq},@Article{rowSeq},@Location{rowSeq},@Line{rowSeq},@FactoryID{rowSeq},@AddDate)
";
                    rowSeq++;
                }
            }

            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }

        #endregion
    }
}
