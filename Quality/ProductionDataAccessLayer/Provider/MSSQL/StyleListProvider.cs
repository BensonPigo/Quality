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


        public int Check_SampleRFTInspection_Count(StyleList_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
            };
            string sqlcmd = $@"
select COUNT(1)
from SampleRFTInspection a WITH(NOLOCK)
inner join SciProduction_Style b on a.StyleUkey = b.Ukey
where b.ID = @StyleID AND b.SeasonID = @SeasonID AND b.BrandID = @BrandID
";

            var result = ExecuteScalar(CommandType.Text, sqlcmd, objParameter);

            return Convert.ToInt32(result == null ? 0 : result);
        }
        public IList<StyleList> Get_StyleInfo(StyleList_Request req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlCmd = string.Empty;

            string RFT_cmd = string.Empty;

            if (req.InspectionTableName== "RFT_Inspection")
            {
                RFT_cmd = $@"
	select Val = ROUND( SUM(IIF(Status = 'Pass',1,0)) * 1.0  / COUNT(1) *1.0  *100 , 2)
    FROM ExtendServer.ManufacturingExecution.dbo.RFT_Inspection  rft WITH(NOLOCK)
	WHERE rft.StyleUkey = s.Ukey
";
            }
            else if (req.InspectionTableName == "SampleRFTInspection")
            {
                RFT_cmd = $@"
	select Val = ROUND( SUM(IIF(Result = 'Pass',1,0)) * 1.0  / COUNT(1) *1.0  *100 , 2)
    FROM ExtendServer.ManufacturingExecution.dbo.SampleRFTInspection  rft WITH(NOLOCK)
	WHERE rft.StyleUkey = s.Ukey AND SubmitDate IS NOT NULL
";
            }

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
	{RFT_cmd}
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
