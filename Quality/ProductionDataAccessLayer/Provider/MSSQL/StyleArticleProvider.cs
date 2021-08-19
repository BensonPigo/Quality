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
    /*Style Article(StyleArticleProvider) 詳細敘述如下*/
    /// <summary>
    /// Style Article
    /// </summary>
    /// <info>Author: Admin; Date: 2021/08/19  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/19  1.00    Admin        Create
    /// </history>
    public class StyleArticleProvider : SQLDAL, IStyleArticleProvider
    {
        #region 底層連線
        public StyleArticleProvider(string ConString) : base(ConString) { }
        public StyleArticleProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base
		/*回傳Style Article(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Style Article
        /// </summary>
        /// <param name="Item">Style Article成員</param>
        /// <returns>回傳Style Article</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public IList<Style_Article> Get(Style_Article Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,TissuePaper"+ Environment.NewLine);
            SbSql.Append("        ,ArticleName"+ Environment.NewLine);
            SbSql.Append("        ,Contents"+ Environment.NewLine);
            SbSql.Append("FROM [Style_Article]"+ Environment.NewLine);



            return ExecuteList<Style_Article>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Style Article(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Style Article
        /// </summary>
        /// <param name="Item">Style Article成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public int Create(Style_Article Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Style_Article]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,Seq"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,TissuePaper"+ Environment.NewLine);
            SbSql.Append("        ,ArticleName"+ Environment.NewLine);
            SbSql.Append("        ,Contents"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @StyleUkey"); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);
            SbSql.Append("        ,@Seq"); objParameter.Add("@Seq", DbType.String, Item.Seq);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@TissuePaper"); objParameter.Add("@TissuePaper", DbType.String, Item.TissuePaper);
            SbSql.Append("        ,@ArticleName"); objParameter.Add("@ArticleName", DbType.String, Item.ArticleName);
            SbSql.Append("        ,@Contents"); objParameter.Add("@Contents", DbType.String, Item.Contents);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Style Article(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Style Article
        /// </summary>
        /// <param name="Item">Style Article成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public int Update(Style_Article Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Style_Article]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.StyleUkey != null) { SbSql.Append("StyleUkey=@StyleUkey"+ Environment.NewLine); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);}
            if (Item.Seq != null) { SbSql.Append(",Seq=@Seq"+ Environment.NewLine); objParameter.Add("@Seq", DbType.String, Item.Seq);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.TissuePaper != null) { SbSql.Append(",TissuePaper=@TissuePaper"+ Environment.NewLine); objParameter.Add("@TissuePaper", DbType.String, Item.TissuePaper);}
            if (Item.ArticleName != null) { SbSql.Append(",ArticleName=@ArticleName"+ Environment.NewLine); objParameter.Add("@ArticleName", DbType.String, Item.ArticleName);}
            if (Item.Contents != null) { SbSql.Append(",Contents=@Contents"+ Environment.NewLine); objParameter.Add("@Contents", DbType.String, Item.Contents);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Style Article(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Style Article
        /// </summary>
        /// <param name="Item">Style Article成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/19  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/19  1.00    Admin        Create
        /// </history>
        public int Delete(Style_Article Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Style_Article]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
