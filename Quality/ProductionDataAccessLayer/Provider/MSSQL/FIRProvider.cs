using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using DatabaseObject.ProductionDB;
using ADOHelper.Utility;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class FIRProvider : SQLDAL, IFIRProvider
    {
        #region 底層連線
        public FIRProvider(string ConString) : base(ConString) { }
        public FIRProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

		#region CRUD Base

	#endregion
    }
}
