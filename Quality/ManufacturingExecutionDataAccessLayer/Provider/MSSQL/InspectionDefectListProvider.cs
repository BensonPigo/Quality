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
	,rfd.DefectPicture
from Orders o 
LEFT JOIN ManufacturingExecution.dbo.RFT_Inspection rf on o.ID = rf.OrderID 
LEFT JOIN ManufacturingExecution.dbo.RFT_Inspection_Detail rfd ON rf.ID = rfd.ID
LEFT JOIN DropdownList d ON d.Type='PMS_RFTResp' AND d.ID = rfd.PMS_RFTRespID
where o.ID = @OrderID
";

            return ExecuteList<InspectionDefectList_Result>(CommandType.Text, sql, listPar).ToList();
        }
    }
}
