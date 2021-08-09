using ADOHelper.Utility.Interface;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace ADOHelper.Utility
{
	public class SQLDataTransaction : ISQLDataTransaction
	{
		private SqlTransaction Transaction;

		private SqlConnection Connection;

		public SqlConnection GetConnection
		{
			get
			{
				return this.Connection;
			}
		}

		public SqlTransaction GetTransaction
		{
			get
			{
				return this.Transaction;
			}
		}

		public SqlConnection SetConnection
		{
			set
			{
				this.Connection = value;
			}
		}

		public SQLDataTransaction(string DataAccessLayer)
		{
			this.Connection = new SqlConnection(ConfigurationManager.ConnectionStrings[DataAccessLayer].ConnectionString);
		}

		public void CloseConnection()
		{
			this.Connection.Close();
		}

		public void Commit()
		{
			this.Transaction.Commit();
		}

		public void OpenConnection()
		{
			if (this.Connection.State == ConnectionState.Closed)
			{
				this.Connection.Open();
				this.Transaction = this.Connection.BeginTransaction();
			}
		}

		public void RollBack()
		{
			this.Transaction.Rollback();
		}
	}
}