using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ProductionDataAccessLayer.Provider
{
    public class PublicWondowProvider : SQLDAL
    {
        #region 底層連線
        public PublicWondowProvider(string ConString) : base(ConString) { }
        public PublicWondowProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<Window_Brand> Get_Brand(string ID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();
            SbSql.Append($@"
select ID
from Production.dbo.Brand
where junk = 0

");
            if (!string.IsNullOrEmpty(ID))
            {

                SbSql.Append($@"AND ID LIKE @ID");
                paras.Add("@ID", DbType.String, ID + "%");
            }

            return ExecuteList<Window_Brand>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Season> Get_Season(string BrandID, string ID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();
            SbSql.Append($@"
select DISTINCT ID
from Production.dbo.Season
where junk = 0

");
            if (!string.IsNullOrEmpty(ID))
            {
                SbSql.Append($@"AND ID LIKE @ID");
                paras.Add("@ID", DbType.String, ID + "%");
            }

            if (!string.IsNullOrEmpty(BrandID))
            {
                SbSql.Append($@"AND BrandID = @BrandID");
                paras.Add("@BrandID", DbType.String, BrandID);
            }

            return ExecuteList<Window_Season>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Style> Get_Style(string BrandID, string SeasonID, string ID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();
            SbSql.Append($@"
select DISTINCT ID
from Production.dbo.Style
where junk = 0

");
            if (!string.IsNullOrEmpty(ID))
            {
                SbSql.Append($@"AND ID LIKE @ID ");
                paras.Add("@ID", DbType.String, ID + "%");
            }

            if (!string.IsNullOrEmpty(BrandID))
            {
                SbSql.Append($@"AND BrandID = @BrandID ");
                paras.Add("@BrandID", DbType.String, BrandID);
            }
            if (!string.IsNullOrEmpty(SeasonID))
            {
                SbSql.Append($@"AND SeasonID = @SeasonID ");
                paras.Add("@SeasonID", DbType.String, SeasonID);
            }

            return ExecuteList<Window_Style>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Article> Get_Article(string OrderID, Int64 StyleUkey, string Article)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            if (!string.IsNullOrEmpty(OrderID))
            {
                //有傳入 OrderID
                SbSql.Append($@"
select DISTINCT Article
from Production.dbo.Order_Qty
where 1=1
AND ID = @OrderID
");
                paras.Add("@OrderID", DbType.String, OrderID);
            }
            else if (string.IsNullOrEmpty(OrderID))
            {
                //沒有傳入 OrderID
                SbSql.Append($@"
select DISTINCT Article
from Production.dbo.Style_Article
where 1=1
AND StyleUkey = @StyleUkey
");
                paras.Add("@StyleUkey ", DbType.Int64, StyleUkey);
            }
            else
            {
                //其餘則回傳空
                return new List<Window_Article>();
            }


            if (!string.IsNullOrEmpty(Article))
            {
                SbSql.Append($@"AND Article LIKE @Article ");
                paras.Add("@Article", DbType.String, Article + "%");
            }

            return ExecuteList<Window_Article>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Article> Get_Size(string OrderID, Int64 StyleUkey, string Article, string Size)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            if (!string.IsNullOrEmpty(OrderID))
            {
                //有傳入 OrderID
                SbSql.Append($@"
select DISTINCT SizeCode
from Production.dbo.Order_Qty
where 1=1
AND ID = @OrderID
");
                paras.Add("@OrderID", DbType.String, OrderID);

                if (!string.IsNullOrEmpty(Article))
                {
                    SbSql.Append($@"AND Article = @Article ");
                    paras.Add("@Article", DbType.String, Article);
                }
            }
            else if (string.IsNullOrEmpty(OrderID))
            {
                //沒有傳入 OrderID
                SbSql.Append($@"
select DISTINCT SizeCode
from Production.dbo.Style_SizeCode
where 1=1
AND StyleUkey = @StyleUkey
");
                paras.Add("@StyleUkey ", DbType.Int64, StyleUkey);
            }
            else
            {
                //其餘則回傳空
                return new List<Window_Article>();
            }


            if (!string.IsNullOrEmpty(Size))
            {
                SbSql.Append($@"AND SizeCode LIKE @Size ");
                paras.Add("@Size", DbType.String, Size + "%");
            }


            return ExecuteList<Window_Article>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Technician> Get_Technician(string CallFunction, string Region, string ID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            SbSql.Append($@"
Select tch.ID
		, p1.Name
		, p1.ExtNo
		, p1.Factory
From Production.dbo.Technician tch
Inner join Production.dbo.Pass1 p1 on tch.ID = p1.ID
Where tch.{CallFunction} = 1

");


            if (!string.IsNullOrEmpty(ID))
            {
                SbSql.Append($@"AND tch.ID LIKE @ID ");
                paras.Add("@ID", DbType.String, ID + "%");
            }

            return ExecuteList<Window_Technician>(CommandType.Text, SbSql.ToString(), paras);
        }

    }
}
