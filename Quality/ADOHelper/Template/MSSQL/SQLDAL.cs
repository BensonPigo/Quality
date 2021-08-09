using ADOHelper.DBToolKit;
using ADOHelper.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace ADOHelper.Template.MSSQL
{
	public abstract class SQLDAL
	{
		private readonly static bool IsSaveErrorLog = false;

		private readonly static bool IsSaveSqlLog = false;


        private static string _ConnectionString;

        //private SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings[_ConnectionString].ConnectionString);
        

        protected SqlDataAdapter adapter;

		protected SqlCommand com;

		static SQLDAL()
		{
            //ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog = (Common.GetAppSetting("IsSaveErrorLog", Assembly.GetExecutingAssembly().GetName().Name) == "Y" ? true : false);
            //ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog = (Common.GetAppSetting("IsSaveSqlLog", Assembly.GetExecutingAssembly().GetName().Name) == "Y" ? true : false);
		}

		public SQLDAL(string ConnectionString)
		{
			try
			{
                _ConnectionString = ConnectionString;
                SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings[_ConnectionString].ConnectionString);
                this.adapter = new SqlDataAdapter(string.Empty, cn);
				this.com = new SqlCommand(string.Empty, cn);
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog)
				{
					//LogHelpe.WriteError(exception.ToString());
				}
				throw exception;
			}
		}

		public SQLDAL(SQLDataTransaction tra)
		{
			try
			{
				tra.OpenConnection();
				this.com = new SqlCommand(string.Empty, tra.GetConnection, tra.GetTransaction);
				this.adapter = new SqlDataAdapter(this.com);
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog)
				{
					//LogHelpe.WriteError(exception.ToString());
				}
				throw exception;
			}
		}

		public T DynamicExecute<T>(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter)
		{
			string name = typeof(T).Name;
			T t = default(T);
			//string str = name;
			//if (str == "DataTable")
			//{
			//	t = (T)this.ExecuteDataTable(CommandType.Text, cmdText, cmdParameter);
			//}
			//else if (str == "String")
			//{
			//	t = (T)this.ExecuteJason(CommandType.Text, cmdText, cmdParameter);
			//}
			//else if (str == "DataSet")
			//{
			//	t = (T)this.ExecuteDataSet(CommandType.Text, cmdText, cmdParameter);
			//}
			//else if (str == "INT")
			//{
			//	t = (T)(object)this.ExecuteNonQuery(CommandType.Text, cmdText, cmdParameter);
			//}
			return t;
		}

		public IList<T> DynamicExecuteList<T>(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter)
		where T : new()
		{
			return this.ExecuteList<T>(CommandType.Text, cmdText, cmdParameter);
		}

		public DataSet ExecuteDataSet(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter, int commTimeout = 30)
		{
			DataSet dataSet;
			try
			{
				DataSet dataSet1 = new DataSet();
				DataTable dataTable = new DataTable();
				string str = cmdText;
				for (int i = 0; i < cmdParameter.Count; i++)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(cmdParameter[i].Value);
					DataAccessHelper.AddParamTOSQLCmd(this.adapter, cmdParameter[i].ParameterName, objectValue, cmdParameter[i].DbType);
					if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
					{
						string parameterName = cmdParameter[i].ParameterName;
						DbType dbType = cmdParameter[i].DbType;
						//LogHelpe.MergerSQLParam(ref str, parameterName, dbType.ToString(), objectValue);
					}
				}
				this.adapter.SelectCommand.CommandText = cmdText;
				this.adapter.SelectCommand.CommandType = cmdType;
				if (commTimeout > 30)
				{
					this.adapter.SelectCommand.CommandTimeout = commTimeout;
				}
				this.adapter.Fill(dataSet1);
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
				{
					//LogHelpe.WriteSQL(cmdText);
				}
				dataSet = dataSet1;
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog)
				{
					//LogHelpe.WriteError(exception.ToString());
				}
				throw exception;
			}
			return dataSet;
		}

		public DataTable ExecuteDataTable(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter)
		{
			DataTable dataTable;
			try
			{
				DataTable dataTable1 = new DataTable();
				string str = cmdText;
				for (int i = 0; i < cmdParameter.Count; i++)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(cmdParameter[i].Value);
					DataAccessHelper.AddParamTOSQLCmd(this.adapter, cmdParameter[i].ParameterName, objectValue, cmdParameter[i].DbType);
					if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
					{
						string parameterName = cmdParameter[i].ParameterName;
						DbType dbType = cmdParameter[i].DbType;
						//LogHelpe.MergerSQLParam(ref str, parameterName, dbType.ToString(), objectValue);
					}
				}
				this.adapter.SelectCommand.CommandText = cmdText;
				this.adapter.SelectCommand.CommandType = cmdType;
				this.adapter.Fill(dataTable1);
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
				{
					//LogHelpe.WriteSQL(cmdText);
				}
				dataTable = dataTable1;
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog)
				{
					//LogHelpe.WriteError(exception.ToString());
				}
				throw exception;
			}
			return dataTable;
		}

		//public string ExecuteJason(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter)
		//{
		//	string str;
		//	try
		//	{
		//		string empty = string.Empty;
		//		string str1 = cmdText;
		//		if (this.com == null)
		//		{
		//			throw new ArgumentNullException("Com");
		//		}
		//		if (this.com.Connection == null)
		//		{
		//			throw new Exception("Connection Is Null");
		//		}
		//		if (this.com.Connection.State == ConnectionState.Closed)
		//		{
		//			this.com.Connection.Open();
		//		}
		//		this.com.Parameters.Clear();
		//		for (int i = 0; i < cmdParameter.Count; i++)
		//		{
		//			object objectValue = RuntimeHelpers.GetObjectValue(cmdParameter[i].Value);
		//			DataAccessHelper.AddParamTOSQLCmd(this.com, cmdParameter[i].ParameterName, objectValue, cmdParameter[i].DbType);
		//			if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
		//			{
		//				string parameterName = cmdParameter[i].ParameterName;
		//				DbType dbType = cmdParameter[i].DbType;
		//				//LogHelpe.MergerSQLParam(ref str1, parameterName, dbType.ToString(), objectValue);
		//			}
		//		}
		//		this.com.CommandType = cmdType;
		//		this.com.CommandText = cmdText;
		//		using (IDataReader dataReader = this.com.ExecuteReader())
		//		{
		//			empty = JsonHelpers.ToJson(dataReader);
		//		}
		//		if (this.com.Transaction == null)
		//		{
		//			this.com.Connection.Close();
		//		}
		//		if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
		//		{
		//			//LogHelpe.WriteSQL(str1);
		//		}
		//		str = empty;
		//	}
		//	catch (Exception exception1)
		//	{
		//		Exception exception = exception1;
		//		if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog)
		//		{
		//			//LogHelpe.WriteError(exception.ToString());
		//		}
		//		throw exception;
		//	}
		//	return str;
		//}

		public IList<T> ExecuteList<T>(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter, int commTimeout = 30)
		where T : new()
		{
			IList<T> ts;
			try
			{
				IList<T> list = new List<T>();
				string str = cmdText;
				if (this.com == null)
				{
					throw new ArgumentNullException("Com");
				}
				if (this.com.Connection == null)
				{
					throw new Exception("Connection Is Null");
				}
				if (this.com.Connection.State == ConnectionState.Closed)
				{
					this.com.Connection.Open();
				}
				this.com.Parameters.Clear();
				for (int i = 0; i < cmdParameter.Count; i++)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(cmdParameter[i].Value);
					DataAccessHelper.AddParamTOSQLCmd(this.com, cmdParameter[i].ParameterName, objectValue, cmdParameter[i].DbType);
					if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
					{
						string parameterName = cmdParameter[i].ParameterName;
						DbType dbType = cmdParameter[i].DbType;
						//LogHelpe.MergerSQLParam(ref str, parameterName, dbType.ToString(), objectValue);
					}
				}
				this.com.CommandType = cmdType;
				this.com.CommandText = cmdText;
                if (commTimeout > 30)
                {
                    this.com.CommandTimeout = commTimeout;
                } 
                using (SqlDataReader sqlDataReader = this.com.ExecuteReader())
				{
					list = DataAccessHelper.DataReaderToList<T>(sqlDataReader);
				}
				if (this.com.Transaction == null)
				{
					this.com.Connection.Close();
				}
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
				{
					//LogHelpe.WriteSQL(str);
				}
				ts = list;
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog)
				{
					//LogHelpe.WriteError(exception.ToString());
				}
				throw exception;
			}
			return ts;
		}

		public int ExecuteNonQuery(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter)
		{
			int num;
			try
			{
				string str = cmdText;
				if (this.com == null)
				{
					throw new ArgumentNullException("Com");
				}
				if (this.com.Connection == null)
				{
					throw new Exception("Connection Is Null");
				}
				if (this.com.Connection.State == ConnectionState.Closed)
				{
					this.com.Connection.Open();
				}
				this.com.Parameters.Clear();
				for (int i = 0; i < cmdParameter.Count; i++)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(cmdParameter[i].Value);
					DataAccessHelper.AddParamTOSQLCmd(this.com, cmdParameter[i].ParameterName, objectValue, cmdParameter[i].DbType);
					if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
					{
						string parameterName = cmdParameter[i].ParameterName;
						DbType dbType = cmdParameter[i].DbType;
						//LogHelpe.MergerSQLParam(ref str, parameterName, dbType.ToString(), objectValue);
					}
				}
				this.com.CommandType = cmdType;
				this.com.CommandText = cmdText;
				int num1 = this.com.ExecuteNonQuery();
				if (this.com.Transaction == null)
				{
					this.com.Connection.Close();
				}
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
				{
					//LogHelpe.WriteSQL(str);
				}
				num = num1;
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog)
				{
					//LogHelpe.WriteError(exception.ToString());
				}
				throw exception;
			}
			return num;
		}

		public int ExecuteNonQuery(CommandType cmdType, string cmdText)
		{
			int num;
			try
			{
				if (this.com == null)
				{
					throw new ArgumentNullException("Com");
				}
				if (this.com.Connection == null)
				{
					throw new Exception("Connection Is Null");
				}
				if (this.com.Connection.State == ConnectionState.Closed)
				{
					this.com.Connection.Open();
				}
				this.com.CommandType = cmdType;
				this.com.CommandText = cmdText;
				int num1 = this.com.ExecuteNonQuery();
				if (this.com.Transaction == null)
				{
					this.com.Connection.Close();
				}
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
				{
					//LogHelpe.WriteSQL(cmdText);
				}
				num = num1;
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog)
				{
					//LogHelpe.WriteError(exception.ToString());
				}
				throw exception;
			}
			return num;
		}

		public object ExecuteScalar(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter)
		{
			object obj;
			try
			{
				string str = cmdText;
				if (this.com == null)
				{
					throw new ArgumentNullException("Com");
				}
				if (this.com.Connection == null)
				{
					throw new Exception("Connection Is Null");
				}
				if (this.com.Connection.State == ConnectionState.Closed)
				{
					this.com.Connection.Open();
				}
				for (int i = 0; i < cmdParameter.Count; i++)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(cmdParameter[i].Value);
					DataAccessHelper.AddParamTOSQLCmd(this.com, cmdParameter[i].ParameterName, objectValue, cmdParameter[i].DbType);
					if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
					{
						string parameterName = cmdParameter[i].ParameterName;
						DbType dbType = cmdParameter[i].DbType;
						//LogHelpe.MergerSQLParam(ref str, parameterName, dbType.ToString(), objectValue);
					}
				}
				this.com.CommandType = cmdType;
				this.com.CommandText = cmdText;
				object obj1 = this.com.ExecuteScalar();
				if (this.com.Transaction == null)
				{
					this.com.Connection.Close();
				}
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
				{
					//LogHelpe.WriteSQL(cmdText);
				}
				obj = obj1;
			}
			catch (Exception exception1)
			{
				Exception exception = exception1;
				if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveErrorLog)
				{
					//LogHelpe.WriteError(exception.ToString());
				}
				throw exception;
			}
			return obj;
		}
	}
}