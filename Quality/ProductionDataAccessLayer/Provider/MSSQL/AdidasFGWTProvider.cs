using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class AdidasFGWTProvider : SQLDAL, IAdidasFGWTProvider
    {
        #region 底層連線
        public AdidasFGWTProvider(string ConString) : base(ConString) { }
        public AdidasFGWTProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
        public IList<Adidas_FGWT> Get(Adidas_FGWT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         Seq"+ Environment.NewLine);
            SbSql.Append("        ,TestName"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,SystemType"+ Environment.NewLine);
            SbSql.Append("        ,ReportType"+ Environment.NewLine);
            SbSql.Append("        ,MtlTypeID"+ Environment.NewLine);
            SbSql.Append("        ,Washing"+ Environment.NewLine);
            SbSql.Append("        ,FabricComposition"+ Environment.NewLine);
            SbSql.Append("        ,TestDetail"+ Environment.NewLine);
            SbSql.Append("        ,Scale"+ Environment.NewLine);
            SbSql.Append("        ,Criteria"+ Environment.NewLine);
            SbSql.Append("        ,Criteria2"+ Environment.NewLine);
            SbSql.Append("FROM [Adidas_FGWT]"+ Environment.NewLine);



            return ExecuteList<Adidas_FGWT>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Create(Adidas_FGWT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Adidas_FGWT]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         Seq"+ Environment.NewLine);
            SbSql.Append("        ,TestName"+ Environment.NewLine);
            SbSql.Append("        ,Location"+ Environment.NewLine);
            SbSql.Append("        ,SystemType"+ Environment.NewLine);
            SbSql.Append("        ,ReportType"+ Environment.NewLine);
            SbSql.Append("        ,MtlTypeID"+ Environment.NewLine);
            SbSql.Append("        ,Washing"+ Environment.NewLine);
            SbSql.Append("        ,FabricComposition"+ Environment.NewLine);
            SbSql.Append("        ,TestDetail"+ Environment.NewLine);
            SbSql.Append("        ,Scale"+ Environment.NewLine);
            SbSql.Append("        ,Criteria"+ Environment.NewLine);
            SbSql.Append("        ,Criteria2"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @Seq"); objParameter.Add("@Seq", DbType.Int32, Item.Seq);
            SbSql.Append("        ,@TestName"); objParameter.Add("@TestName", DbType.String, Item.TestName);
            SbSql.Append("        ,@Location"); objParameter.Add("@Location", DbType.String, Item.Location);
            SbSql.Append("        ,@SystemType"); objParameter.Add("@SystemType", DbType.String, Item.SystemType);
            SbSql.Append("        ,@ReportType"); objParameter.Add("@ReportType", DbType.String, Item.ReportType);
            SbSql.Append("        ,@MtlTypeID"); objParameter.Add("@MtlTypeID", DbType.String, Item.MtlTypeID);
            SbSql.Append("        ,@Washing"); objParameter.Add("@Washing", DbType.String, Item.Washing);
            SbSql.Append("        ,@FabricComposition"); objParameter.Add("@FabricComposition", DbType.String, Item.FabricComposition);
            SbSql.Append("        ,@TestDetail"); objParameter.Add("@TestDetail", DbType.String, Item.TestDetail);
            SbSql.Append("        ,@Scale"); objParameter.Add("@Scale", DbType.String, Item.Scale);
            SbSql.Append("        ,@Criteria"); objParameter.Add("@Criteria", DbType.String, Item.Criteria);
            SbSql.Append("        ,@Criteria2"); objParameter.Add("@Criteria2", DbType.String, Item.Criteria2);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update(Adidas_FGWT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Adidas_FGWT]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.Seq != null) { SbSql.Append("Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.Int32, Item.Seq);}
            if (Item.TestName != null) { SbSql.Append(",TestName=@TestName"+ Environment.NewLine); objParameter.Add("@TestName", DbType.String, Item.TestName);}
            if (Item.Location != null) { SbSql.Append(",Location=@Location"+ Environment.NewLine); objParameter.Add("@Location", DbType.String, Item.Location);}
            if (Item.SystemType != null) { SbSql.Append(",SystemType=@SystemType"+ Environment.NewLine); objParameter.Add("@SystemType", DbType.String, Item.SystemType);}
            if (Item.ReportType != null) { SbSql.Append(",ReportType=@ReportType"+ Environment.NewLine); objParameter.Add("@ReportType", DbType.String, Item.ReportType);}
            if (Item.MtlTypeID != null) { SbSql.Append(",MtlTypeID=@MtlTypeID"+ Environment.NewLine); objParameter.Add("@MtlTypeID", DbType.String, Item.MtlTypeID);}
            if (Item.Washing != null) { SbSql.Append(",Washing=@Washing"+ Environment.NewLine); objParameter.Add("@Washing", DbType.String, Item.Washing);}
            if (Item.FabricComposition != null) { SbSql.Append(",FabricComposition=@FabricComposition"+ Environment.NewLine); objParameter.Add("@FabricComposition", DbType.String, Item.FabricComposition);}
            if (Item.TestDetail != null) { SbSql.Append(",TestDetail=@TestDetail"+ Environment.NewLine); objParameter.Add("@TestDetail", DbType.String, Item.TestDetail);}
            if (Item.Scale != null) { SbSql.Append(",Scale=@Scale"+ Environment.NewLine); objParameter.Add("@Scale", DbType.String, Item.Scale);}
            if (Item.Criteria != null) { SbSql.Append(",Criteria=@Criteria"+ Environment.NewLine); objParameter.Add("@Criteria", DbType.String, Item.Criteria);}
            if (Item.Criteria2 != null) { SbSql.Append(",Criteria2=@Criteria2"+ Environment.NewLine); objParameter.Add("@Criteria2", DbType.String, Item.Criteria2);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Delete(Adidas_FGWT Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Adidas_FGWT]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
