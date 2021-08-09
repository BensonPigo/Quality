using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.CompilerServices;

namespace ADOHelper.Template.MSSQL
{
	public class SQLParameterCollection : List<SqlParameter>
	{
		public SQLParameterCollection()
		{
		}

		public void Add(string ParameterName, object Value)
		{
			base.Add(new SqlParameter()
			{
				ParameterName = ParameterName,
				Value = RuntimeHelpers.GetObjectValue(Value)
			});
		}

		public void Add(string ParameterName, System.Data.DbType DbType, object Value)
		{
			base.Add(new SqlParameter()
			{
				ParameterName = ParameterName,
				DbType = DbType,
				Value = RuntimeHelpers.GetObjectValue(Value)
			});
		}

		public void Add(string ParameterName, System.Data.DbType DbType, object Value, System.Data.ParameterDirection ParameterDirection, int Size)
		{
			base.Add(new SqlParameter()
			{
				ParameterName = ParameterName,
				DbType = DbType,
				Value = RuntimeHelpers.GetObjectValue(Value),
				Direction = ParameterDirection,
				Size = Size
			});
		}
	}
}