using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    /*Style - Bill of Accessory(StyleBOAProvider) 詳細敘述如下*/
    /// <summary>
    /// Style - Bill of Accessory
    /// </summary>
    /// <info>Author: Admin; Date: 2021/09/03  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/09/03  1.00    Admin        Create
    /// </history>
    public class StyleBOAProvider : SQLDAL, IStyleBOAProvider
    {
        #region 底層連線
        public StyleBOAProvider(string ConString) : base(ConString) { }
        public StyleBOAProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base

        public IList<Style_BOA> Get(Style_BOA Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         StyleUkey"+ Environment.NewLine);
            SbSql.Append("        ,Ukey"+ Environment.NewLine);
            SbSql.Append("        ,Refno"+ Environment.NewLine);
            SbSql.Append("        ,SCIRefno"+ Environment.NewLine);
            SbSql.Append("        ,SEQ1"+ Environment.NewLine);
            SbSql.Append("        ,ConsPC"+ Environment.NewLine);
            SbSql.Append("        ,PatternPanel"+ Environment.NewLine);
            SbSql.Append("        ,SizeItem"+ Environment.NewLine);
            SbSql.Append("        ,ProvidedPatternRoom"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,ColorDetail"+ Environment.NewLine);
            SbSql.Append("        ,IsCustCD"+ Environment.NewLine);
            SbSql.Append("        ,BomTypeZipper"+ Environment.NewLine);
            SbSql.Append("        ,BomTypeSize"+ Environment.NewLine);
            SbSql.Append("        ,BomTypeColor"+ Environment.NewLine);
            SbSql.Append("        ,BomTypeStyle"+ Environment.NewLine);
            SbSql.Append("        ,BomTypeArticle"+ Environment.NewLine);
            SbSql.Append("        ,BomTypePo"+ Environment.NewLine);
            SbSql.Append("        ,BomTypeCustCD"+ Environment.NewLine);
            SbSql.Append("        ,BomTypeFactory"+ Environment.NewLine);
            SbSql.Append("        ,BomTypeBuyMonth"+ Environment.NewLine);
            SbSql.Append("        ,BomTypeCountry"+ Environment.NewLine);
            SbSql.Append("        ,SuppIDBulk"+ Environment.NewLine);
            SbSql.Append("        ,SuppIDSample"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,FabricPanelCode"+ Environment.NewLine);
            SbSql.Append("FROM [Style_BOA] WITH(NOLOCK)" + Environment.NewLine);

            return ExecuteList<Style_BOA>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<Style_BOA> GetAccessoryRefNo(AccessoryRefNo_Request Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append(@"
SELECT Refno = ''
UNION
select sb.Refno
from Style_BOA sb WITH(NOLOCK)
inner join Fabric f WITH(NOLOCK) on sb.SCIRefno = f.SCIRefno
Where 1 = 1
");
            if (!string.IsNullOrEmpty(Item.MtlTypeID))
            {
                SbSql.Append("And f.MtlTypeID = @MtlTypeID" + Environment.NewLine);
                objParameter.Add("@MtlTypeID", DbType.String, Item.MtlTypeID);
            }

            if (Item.StyleUkey != null)
            {
                SbSql.Append("And sb.StyleUkey = @StyleUkey" + Environment.NewLine);
                objParameter.Add("@StyleUkey",  Item.StyleUkey);
            }

            if (!string.IsNullOrEmpty(Item.BrandID) && !string.IsNullOrEmpty(Item.SeasonID) && !string.IsNullOrEmpty(Item.StyleID))
            {
                SbSql.Append("And sb.StyleUkey = (select ukey from Style WITH(NOLOCK) where ID = @StyleID and SeasonID = @SeasonID and BrandID = @BrandID)" + Environment.NewLine);
                objParameter.Add("@StyleID", DbType.String, Item.StyleID);
                objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
                objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            }

            SbSql.Append("Order by Refno");

            return ExecuteList<Style_BOA>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion
    }
}
