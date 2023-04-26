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

		public int Check_SampleRFTInspection_Count(string OrderID)
		{
			SQLParameterCollection objParameter = new SQLParameterCollection
			{
				{ "@OrderID", DbType.String, OrderID } ,
			};
			string sqlcmd = $@"
select COUNT(1)
from SampleRFTInspection WITH(NOLOCK)
where OrderID = @OrderID
";

			var result = ExecuteScalar(CommandType.Text, sqlcmd, objParameter);

			return Convert.ToInt32(result == null ? 0 : result);
		}

		public IList<InspectionDefectList_Result> GetData(string OrderID)
        {

            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add(new SqlParameter("@OrderID", OrderID));

            string sql = $@"
select rfd.*
INTO #Inspection_Detail
from ExtendServer.ManufacturingExecution.dbo.RFT_Inspection rf WITH(NOLOCK) 
INNER JOIN ExtendServer.ManufacturingExecution.dbo.RFT_Inspection_Detail rfd WITH(NOLOCK) ON rf.ID = rfd.ID
where rf.OrderID  = @OrderID

select  DefectType = gdt.ID +'-'+gdt.Description

	,DefectCode = gdc.ID +'-'+gdc.Description

	--,AreaCodes = ISNULL( Areas.Val ,'')
	,AreaCodes = b.AreaCode
	,pmsFile.DefectPicture
from GarmentDefectType gdt 
inner join GarmentDefectCode gdc on gdt.id=gdc.GarmentDefectTypeID 
INNER JOIN #Inspection_Detail b ON  b.GarmentDefectTypeID=gdt.ID AND b.GarmentDefectCodeID=gdc.ID
LEFT JOIN SciPMSFile_RFT_Inspection_Detail pmsFile WITH(NOLOCK) ON pmsFile.Ukey = b.Ukey
where 1=1 and gdt.Junk =0 and gdc.Junk = 0
order by gdt.id,gdc.id

Drop table #Inspection_Detail
";

            return ExecuteList<InspectionDefectList_Result>(CommandType.Text, sql, listPar).ToList();
		}
		public IList<InspectionDefectList_Result> GetData_SampleRFTInspection(string OrderID)
		{

			SQLParameterCollection listPar = new SQLParameterCollection();
			listPar.Add(new SqlParameter("@OrderID", OrderID));

			string sql = $@"

select rfd.*
INTO #Inspection_Detail
from SampleRFTInspection rf WITH(NOLOCK) 
INNER JOIN SampleRFTInspection_Detail rfd WITH(NOLOCK) ON rf.ID = rfd.SampleRFTInspectionUkey
where rf.OrderID  = @OrderID and rfd.Qty > 0

select  DefectType = gdt.ID +'-'+gdt.Description

	,DefectCode = gdc.ID +'-'+gdc.Description

	,AreaCodes = ISNULL( Areas.Val ,'')
	,HasImage =  CAST( IIF( EXISTS(
	
		select 1
		from PMSFile.dbo.SampleRFTInspection_Detail
		where SampleRFTInspectionDetailUKey IN (
			select top 1 Ukey
			from #Inspection_Detail a
			where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
		)
	), 1,0) as bit )
from MainServer.Production.dbo.GarmentDefectType gdt 
inner join MainServer.Production.dbo.GarmentDefectCode gdc on gdt.id=gdc.GarmentDefectTypeID 
outer apply(

	select Val = stuff((
		select  ',' +a.AreaCode
		from #Inspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
		FOR XML PATH('')
		),1,1,'')
)Areas
where 1=1 and gdt.Junk =0 and gdc.Junk =0　AND　Areas.Val IS not null
order by gdt.id,gdc.id

Drop table #Inspection_Detail
";

			return ExecuteList<InspectionDefectList_Result>(CommandType.Text, sql, listPar).ToList();
		}
	}
}
