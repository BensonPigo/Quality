using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using ProductionDataAccessLayer.Interface;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class StyleListProvider : SQLDAL, IStyleListProvider
    {
        #region 底層連線
        public StyleListProvider(string ConString) : base(ConString) { }
        public StyleListProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public IList<StyleList> Get_StyleInfo(StyleList_Request req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlCmd = string.Empty;

            string RFT_cmd = string.Empty;


            sqlCmd = $@"

select StyleID = s.ID
	,s.BrandID
	,s.SeasonID
	,s.Description
	,s.Phase
	,SMR = SMR.Val
	,Handle = Handle.Val
	,SpecialMark = r.Name
    ,SampleRFT = LEFT( CAST(RFT.Val as varchar),5)
from Style s WITH(NOLOCK)
LEFT JOIN Reason r WITH(NOLOCK) ON  r.ReasonTypeID = 'Style_SpecialMark'  and r.ID = s.SpecialMark
OUTER APPLY(
	SELECT Val = p.ID +'-' +p.Name
	FROM TPEPass1  p WITH(NOLOCK)
	WHERE p.ID = s.BulkSMR AND s.Phase = 'Bulk'
	UNION
	SELECT Val = p.ID +'-' +p.Name
	FROM TPEPass1  p WITH(NOLOCK)
	WHERE p.ID = s.SampleSMR AND s.Phase = 'Sample'
)SMR
OUTER APPLY(
	SELECT Val = p.ID +'-' +p.Name
	FROM TPEPass1  p WITH(NOLOCK)
	WHERE p.ID = s.BulkMRHandle AND s.Phase = 'Bulk'
	UNION
	SELECT Val = p.ID +'-' +p.Name
	FROM TPEPass1  p WITH(NOLOCK)
	WHERE p.ID = s.SampleMRHandle AND s.Phase = 'Sample'
)Handle
OUTER APPLY(
	select Val = ROUND( SUM(IIF(Status = 'Pass',1,0)) * 1.0  / COUNT(1) *1.0  *100 , 2)
    FROM ExtendServer.ManufacturingExecution.dbo.RFT_Inspection  rft WITH(NOLOCK)
	WHERE rft.StyleUkey = s.Ukey
)RFT
WHERE 1=1
";

            if (!string.IsNullOrEmpty(req.StyleID))
            {
                sqlCmd += " and s.ID = @StyleID";
                listPar.Add("@StyleID", DbType.String, req.StyleID);
            }

            if (!string.IsNullOrEmpty(req.BrandID))
            {
                sqlCmd += " and s.BrandID = @BrandID";
                listPar.Add("@BrandID", DbType.String, req.BrandID);
            }

            if (!string.IsNullOrEmpty(req.SeasonID))
            {
                sqlCmd += " and s.SeasonID = @SeasonID";
                listPar.Add("@SeasonID", DbType.String, req.SeasonID);
            }

            return ExecuteList<StyleList>(CommandType.Text, sqlCmd, listPar);
        }
    }
}
