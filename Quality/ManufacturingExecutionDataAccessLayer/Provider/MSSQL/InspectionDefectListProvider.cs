using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DatabaseObject.ViewModel.SampleRFT;
using System.Data;
using System.Data.SqlClient;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class InspectionDefectListProvider : SQLDAL
    {
        #region 底層連線
        public InspectionDefectListProvider(string ConString) : base(ConString) { }
        public InspectionDefectListProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<InspectionDefectList_Result> GetData(string OrderID)
        {

            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add(new SqlParameter("@OrderID", OrderID));

            string sql = $@"
Select OrderID = o.ID
	,SampleStage = o.OrderTypeID
	,o.StyleID
	,o.SeasonID
	,rf.Article
	,rf.Size
	,rf.InspectionDate
	,DefectType = Concat (rfd.AreaCode,  ' -', rfd.DefectCode, ' -', d.Description, ' -', rfd.PMS_RFTBACriteriaID)
	,pmsFile.DefectPicture
from Orders o  WITH(NOLOCK)
INNER JOIN ExtendServer.ManufacturingExecution.dbo.RFT_Inspection rf WITH(NOLOCK) on o.ID = rf.OrderID 
INNER JOIN ExtendServer.ManufacturingExecution.dbo.RFT_Inspection_Detail rfd WITH(NOLOCK) ON rf.ID = rfd.ID
LEFT JOIN SciPMSFile_RFT_Inspection_Detail pmsFile WITH(NOLOCK) ON pmsFile.Ukey = rfd.Ukey
LEFT JOIN DropdownList d WITH(NOLOCK) ON d.Type='PMS_RFTResp' AND d.ID = rfd.PMS_RFTRespID
where o.ID = @OrderID
";

            return ExecuteList<InspectionDefectList_Result>(CommandType.Text, sql, listPar).ToList();
        }
    }
}
