using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;

namespace ADOHelper.Template.MSSQL
{
	public class DataAccessHelper
	{
		public DataAccessHelper()
		{
		}

		public static void AddParamTOSQLCmd(SqlDataAdapter adapter, string ParamID, object ParamValue, DbType dbType)
		{
			DataAccessHelper.AddParamTOSQLCmd(adapter.SelectCommand, ParamID, ParamValue, dbType);
		}

		public static void AddParamTOSQLCmd(SqlCommand Com, string ParamID, object ParamValue, DbType dbType)
		{
			DataAccessHelper.AddParamTOSQLCmd(Com, ParamID, ParamValue);
			Com.Parameters[ParamID].DbType = dbType;
		}

		public static void AddParamTOSQLCmd(SqlDataAdapter adapter, string ParamID, object ParamValue)
		{
			DataAccessHelper.AddParamTOSQLCmd(adapter.SelectCommand, ParamID, ParamValue);
		}

		public static void AddParamTOSQLCmd(SqlCommand Com, string ParamID, object ParamValue)
		{
			if (Com == null)
			{
				throw new ArgumentNullException("Com");
			}
			if (string.IsNullOrEmpty(ParamID))
			{
				throw new ArgumentOutOfRangeException("ParamID");
			}
			if (Com.Parameters.IndexOf(ParamID) != -1)
			{
				Com.Parameters.Remove(Com.Parameters[ParamID]);
			}
			if (ParamValue == null)
			{
				Com.Parameters.AddWithValue(ParamID, DBNull.Value);
			}
			else if ((!(ParamValue is DateTime) ? false : (DateTime)ParamValue == DateTime.MinValue))
			{
				Com.Parameters.AddWithValue(ParamID, DBNull.Value);
			}
			else if ((!(ParamValue is string) ? false : ParamValue.ToString() == string.Empty))
			{
				Com.Parameters.AddWithValue(ParamID, "");
			}
			else if ((!(ParamValue is byte[]) ? true : ((byte[])ParamValue).Length != 0))
			{
				Com.Parameters.AddWithValue(ParamID, ParamValue);
			}
			else
			{
				Com.Parameters.Add(ParamID, SqlDbType.Binary);
				Com.Parameters[ParamID].Value = DBNull.Value;
			}
		}

		public static int ConvertToInt(object value)
		{
			int num;
			int num1;
			num1 = (!int.TryParse(value.ToString(), out num) ? -2147483648 : num);
			return num1;
		}

		public static IList<T> DataReaderToList<T>(SqlDataReader rdr)
		{
			IList<T> ts = new List<T>();
			while (rdr.Read())
			{
				T t = Activator.CreateInstance<T>();
				Type type = t.GetType();
				for (int i = 0; i < rdr.FieldCount; i++)
				{
					PropertyInfo property = type.GetProperty(rdr.GetName(i));
					if (property != null)
					{
						object obj = null;
						obj = (!rdr.IsDBNull(i) ? rdr.GetValue(i) : DataAccessHelper.GetDBNullValue(property.PropertyType.FullName));
						property.SetValue(t, obj, null);
					}
				}
				ts.Add(t);
			}
			return ts;
		}

		public static object DataReaderToObj<T>(SqlDataReader rdr)
		{
			object obj;
			T t = Activator.CreateInstance<T>();
			Type type = t.GetType();
			if (!rdr.Read())
			{
				obj = null;
			}
			else
			{
				for (int i = 0; i < rdr.FieldCount; i++)
				{
					object value = null;
					if (!rdr.IsDBNull(i))
					{
						value = rdr.GetValue(i);
					}
					else
					{
						string fullName = type.GetProperty(rdr.GetName(i)).PropertyType.FullName;
						value = DataAccessHelper.GetDBNullValue(fullName);
					}
					type.GetProperty(rdr.GetName(i)).SetValue(t, value, null);
				}
				obj = t;
			}
			return obj;
		}

		public static CommandType GetCommandType(string commandType)
		{
			CommandType commandType1;
			string str = commandType;
			if (str == "T-SQL")
			{
				commandType1 = CommandType.Text;
			}
			else
			{
				commandType1 = (str == "SP" ? CommandType.StoredProcedure : CommandType.Text);
			}
			return commandType1;
		}

		private static object GetDBNullValue(string typeFullName)
		{
			return null;
		}

		public static DbType GetDBType(string dbType)
		{
			DbType dbType1;
			string str = dbType;
			if (str == "String")
			{
				dbType1 = DbType.String;
			}
			else if (str == "Date")
			{
				dbType1 = DbType.Date;
			}
			else
			{
				dbType1 = (str == "DateTime" ? DbType.DateTime : DbType.String);
			}
			return dbType1;
		}
	}
}