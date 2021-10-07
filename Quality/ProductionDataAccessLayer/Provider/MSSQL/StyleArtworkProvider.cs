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
    public class StyleArtworkProvider : SQLDAL, IStyleArtworkProvider
    {
        #region 底層連線
        public StyleArtworkProvider(string ConString) : base(ConString) { }
        public StyleArtworkProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base

        public IList<Style_Artwork> GetArtworkTypeID(StyleArtwork_Request Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append(@"
select distinct ArtworkTypeID
from Style_Artwork
Where 1 = 1
");

            if (!string.IsNullOrEmpty(Item.BrandID) && !string.IsNullOrEmpty(Item.SeasonID) && !string.IsNullOrEmpty(Item.StyleID))
            {
                SbSql.Append("And StyleUkey = (select ukey from Style where ID = @StyleID and SeasonID = @SeasonID and BrandID = @BrandID)" + Environment.NewLine);
                objParameter.Add("@StyleID", DbType.String, Item.StyleID);
                objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
                objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            }

            if (!string.IsNullOrEmpty(Item.POID))
            {
                SbSql.Append("AND StyleUkey = (SELECT StyleUkey From Orders o  WHERE o.ID = @POID )" + Environment.NewLine);
                objParameter.Add("@POID", DbType.String, Item.POID);
            }

            SbSql.Append("order by ArtworkTypeID");

            return ExecuteList<Style_Artwork>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion
    }
}
