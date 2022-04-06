using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using ProductionDataAccessLayer.Interface;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class InspectionTypeProvider : SQLDAL, IInspectionTypeProvider
    {
        public InspectionTypeProvider(string ConString) : base(ConString) { }
        public InspectionTypeProvider(SQLDataTransaction tra) : base(tra) { }

        public IList<InspectionType> Get_InspectionType(string Function, string Category, string BrandID = "ADIDAS")
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            BrandID = string.IsNullOrEmpty(BrandID) || BrandID != "REEBOK" ? "ADIDAS" : BrandID;
            SbSql.Append(
                $@"
select i.* 
from InspectionType i
left join InspectionType_Detail id on i.ID = id.InspectionTypeID
where i.Category = '{Category}'
and id.[Function] = '{Function}'
and id.BrandID = '{BrandID}'
" + Environment.NewLine);

            return ExecuteList<InspectionType>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
