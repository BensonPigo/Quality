using ADOHelper.DBToolKit;
using ADOHelper.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
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

        public DataTable ExecuteDataTableByServiceConn(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter, int commTimeout = 30)
        {
            DataTable dataTable = new DataTable();
            try
            {
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
                this.adapter.Fill(dataTable);
                if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
                {
                    //LogHelpe.WriteSQL(cmdText);
                }
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

		public static DataTable ExecuteDataTable(CommandType cmdType, string cmdText, SQLParameterCollection cmdParameter, string ConnectionString = "")
		{
			DataTable dataTable;
			try
			{
				DataTable dataTable1 = new DataTable();
				string str = cmdText;
                SqlDataAdapter adapter = new SqlDataAdapter();
				ConnectionString = string.IsNullOrEmpty(ConnectionString) ? _ConnectionString : ConnectionString;
				SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings[ConnectionString].ConnectionString);
                adapter = new SqlDataAdapter(string.Empty, cn);
                for (int i = 0; i < cmdParameter.Count; i++)
				{
					object objectValue = RuntimeHelpers.GetObjectValue(cmdParameter[i].Value);
					DataAccessHelper.AddParamTOSQLCmd(adapter, cmdParameter[i].ParameterName, objectValue, cmdParameter[i].DbType);
					if (ADOHelper.Template.MSSQL.SQLDAL.IsSaveSqlLog)
					{
						string parameterName = cmdParameter[i].ParameterName;
						DbType dbType = cmdParameter[i].DbType;
						//LogHelpe.MergerSQLParam(ref str, parameterName, dbType.ToString(), objectValue);
					}
				}
				adapter.SelectCommand.CommandText = cmdText;
				adapter.SelectCommand.CommandType = cmdType;
				adapter.Fill(dataTable1);
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

		public int ExecuteNonQuery(CommandType cmdType, string cmdText, int timeout = 30)
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
				this.com.CommandTimeout = timeout;
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

        /// <summary>
        /// GetID()
        /// </summary>
        /// <param name="string keyWord"></param>
        /// <param name="string tableName"></param>
        /// <param name="[DateTime refDate = Today]"></param>
        /// <param name="[int format = 2]"></param>
        /// <param name="[string checkColumn = "ID"]"></param>
        /// <param name="[string connectionName = null]"></param>
        /// <param name="[int sequenceMode = 1]"></param>
        /// <param name="[int sequenceLength = 0]"></param>
        /// <param name="addNum">[int addNum = 0] 指定ID後第幾筆，只支援sequenceMode=1 </param>
        /// <returns></returns>
        public string GetID(string keyWord, string tableName, DateTime refDate = default(DateTime), int format = 2, string checkColumn = "ID", int sequenceMode = 1, int sequenceLength = 0, int addNum = 1)
        {
            if (String.IsNullOrWhiteSpace(tableName))
            {
                throw new Exception("Parameter - tableName is not specified.."); ;
            }

            if (refDate == DateTime.MinValue)
            {
                refDate = DateTime.Today;
            }

            string TaiwanYear;
            switch (format)
            {
                case 1:     // A yy xxxx
                    keyWord = keyWord.ToUpper().Trim() + refDate.ToString("yy");
                    break;
                case 2:     // A yyMM xxxx
                    keyWord = keyWord.ToUpper().Trim() + refDate.ToString("yyMM");
                    break;
                case 3:      // A yyMMdd xxxx
                    keyWord = keyWord.ToUpper().Trim() + refDate.ToString("yyMMdd");
                    break;
                case 4:      // A yyyyMM xxxxx
                    keyWord = keyWord.ToUpper().Trim() + refDate.ToString("yyyyMM");
                    break;
                case 5:     // 民國年 A yyyMM xxxx
                    TaiwanYear = ((refDate.Year - 1911).ToString()).PadLeft(3, '0');
                    keyWord = keyWord.ToUpper().Trim() + TaiwanYear + refDate.ToString("MM");
                    break;
                case 6:     // A xxxx
                    keyWord = keyWord.ToUpper().Trim();
                    break;
                case 7:    // A yyyy xxxx
                    keyWord = keyWord.ToUpper().Trim() + refDate.ToString("yyyy");
                    break;
                case 8:    // 民國年 A yyyMMdd xxxx
                    TaiwanYear = ((refDate.Year - 1911).ToString()).PadLeft(3, '0');
                    keyWord = keyWord.ToUpper().Trim() + TaiwanYear + refDate.ToString("MM") + refDate.ToString("dd");
                    break;
                default:
                    throw new Exception("Parameter - formatting is incorrect or not found!");
            }

            //判斷schema欄位的結構長度
            string returnID = "";


            string sqlCmd = string.Format("SELECT TOP 1 {0} FROM {1} WHERE {2} LIKE '{3}%' ORDER BY {4} DESC", checkColumn, tableName, checkColumn, keyWord.Trim(), checkColumn);

            DataTable dtID = ExecuteDataTableByServiceConn(CommandType.Text, sqlCmd, new SQLParameterCollection());

			sqlCmd = $@"
select s.max_length
from sys.tables t
inner join sys.columns s on s.object_id = t.object_id
where t.name = '{tableName}'
and s.name = '{checkColumn}'
";

			DataTable dtKey = ExecuteDataTableByServiceConn(CommandType.Text, sqlCmd, new SQLParameterCollection());
			int columnTypeLength = int.Parse(dtKey.Rows[0][0].ToString());

			if (dtID.Rows.Count > 0)
            {
                string lastID = dtID.Rows[0][checkColumn].ToString();
                returnID = keyWord + GetNextValue(lastID.Substring(keyWord.Length), sequenceMode, addNum);
            }
            else
            {
                if (sequenceLength > 0)
                {
                    if ((columnTypeLength - keyWord.Length) >= sequenceLength)
                    {
                        returnID = keyWord + GetNextValue("0".PadLeft(sequenceLength, '0'), sequenceMode, addNum);
                    }
                    else
                    {
                        throw new Exception("(columnTypeLength - keyWord.Length) < sequenceLength"
                        + Environment.NewLine + "columnTypeLength = " + columnTypeLength
                        + Environment.NewLine + "keyWord.Length = " + keyWord.Length
                        + Environment.NewLine + "sequenceLength = " + sequenceLength);
                    }
                }
                else
                {
                    returnID = keyWord + GetNextValue("0".PadLeft(columnTypeLength - keyWord.Length, '0'), sequenceMode);
                }
            }


            return returnID;
        }

        public string GetNextValue(string strValue, int sequenceMode, int addNum = 1)
        {
            char[] charValue = strValue.ToArray<char>();
            int sequenceValue = 0;
            string returnValue = "";
            int charAscii = 0;

            if (sequenceMode == 1)
            {
                // 當第一個字為字母
                if (System.Convert.ToInt32(charValue[0]) >= 65 && System.Convert.ToInt32(charValue[0]) <= 90)
                {
                    sequenceValue = System.Convert.ToInt32(strValue.Substring(1));
                    // 進位處理
                    if ((sequenceValue + addNum).ToString().Length > sequenceValue.ToString().Length && (sequenceValue + addNum).ToString().Length > strValue.Substring(1).Length)
                    {
                        charAscii = System.Convert.ToInt32(charValue[0]);
                        if (charAscii + 1 > 90)
                        {
                            return strValue;
                        }
                        else
                        {
                            // B9751 +25000 =>e4752
                            int num = System.Convert.ToInt32((sequenceValue + addNum + 1).ToString().Substring(1));

                            charValue[0] = System.Convert.ToChar(CarryLetter(charAscii, CarryNum(strValue, sequenceValue + addNum + 1)));
                            sequenceValue = num;

                        }
                    }
                    else
                    {
                        sequenceValue = sequenceValue + addNum;
                    }
                    returnValue = charValue[0] + sequenceValue.ToString().PadLeft(strValue.Length - 1, '0');
                }
                else
                {
                    sequenceValue = System.Convert.ToInt32(strValue);

                    // 進位處理
                    if (((sequenceValue + addNum).ToString().Length > sequenceValue.ToString().Length) && (sequenceValue + addNum).ToString().Length > strValue.Length)
                    {
                        int numSeq = 0;

                        if ((sequenceValue + addNum + 1).ToString().Length > strValue.ToString().Length)
                        {
                            // 89751 +25000 => B4752
                            numSeq = System.Convert.ToInt32((sequenceValue + addNum + 1).ToString().Substring((sequenceValue + addNum + 1).ToString().Length - strValue.ToString().Length + 1));
                        }
                        else
                        {
                            // 99998+5 =A00004
                            numSeq = System.Convert.ToInt32((sequenceValue + addNum + 1).ToString().Substring(1));
                        }

                        charValue[0] = System.Convert.ToChar(CarryLetter(System.Convert.ToChar(strValue.Substring(0, 1)), CarryNum(strValue, sequenceValue + addNum + 1)));
                        returnValue = charValue[0] + numSeq.ToString().PadLeft(strValue.Length - 1, '0');
                    }
                    else
                    {
                        sequenceValue = sequenceValue + addNum;
                        returnValue = sequenceValue.ToString().PadLeft(strValue.Length, '0');
                    }
                }
            }
            else
            {
                for (int i = charValue.Length - 1; i >= 0; i--)
                {
                    charAscii = System.Convert.ToInt32(charValue[i]);

                    if (charAscii == 57)   // 遇9跳A
                    {
                        charValue[i] = 'A';
                        break;
                    }

                    if (charAscii == 72 || charAscii == 78) // I or O略過
                    {
                        charValue[i] = System.Convert.ToChar(charAscii + 2);
                        break;
                    }

                    if (charAscii == 90)  //當字母為Z
                    {
                        if (i > 0)
                        {
                            charValue[i] = '0';
                            continue;
                        }
                        else
                        {
                            return strValue;    //超出最大上限ZZZ...., 返回原值
                        }
                    }

                    charValue[i] = System.Convert.ToChar(charAscii + 1);
                    break;
                }
                returnValue = new String(charValue);
            }
            return returnValue;
        }

        /// <summary>
        /// 進位後傳還新的字母
        /// </summary>
        /// <param name="startAscii">startAscii</param>
        /// <param name="carryTimes">carryTimes</param>
        /// <returns>CarryLetter</returns>
        public static int CarryLetter(int startAscii, int carryTimes)
        {
            // 48~57 -->0~9
            // 65~90 -->A~Z
            int newAscii = startAscii;
            if (!((startAscii >= 48 && startAscii <= 57) || (startAscii >= 65 && startAscii <= 90)))
            {
                return newAscii;
            }

            for (int i = 0; i < carryTimes; i++)
            {
                if (newAscii < 57)
                {
                    newAscii += 1;
                }
                else if (newAscii == 57)
                {
                    newAscii = 65;
                }
                else
                {
                    if (newAscii == 72 || newAscii == 78)
                    {
                        newAscii += 2;
                    }
                    else
                    {
                        newAscii += 1;
                    }
                }
            }

            return newAscii;
        }

        /// <summary>
        /// 進位次數
        /// </summary>
        /// <param name="original">original</param>
        /// <param name="newNum">newNum</param>
        /// <returns>CarryNum</returns>
        public static int CarryNum(string strValue, decimal newNum)
        {
            int tempOldValue = 0;
            int tempNewValue = 0;

            if (strValue.ToString().Length == newNum.ToString().Length)
            {
                if (System.Convert.ToInt32(System.Convert.ToChar(strValue.Substring(0, 1))) >= 65 && System.Convert.ToInt32(System.Convert.ToChar(strValue.Substring(0, 1))) <= 90)
                {
                    return System.Convert.ToInt32(newNum.ToString().Substring(0, 1));
                }
                else
                {
                    return System.Convert.ToInt16(newNum.ToString().Substring(0, 1)) - System.Convert.ToInt16(strValue.ToString().Substring(0, 1));
                }
            }
            else if (int.TryParse(strValue, out tempOldValue) && (int.TryParse(strValue, out tempNewValue) && newNum.ToString().Length > strValue.ToString().Length))
            {
                int carryNum = newNum.ToString().Length - strValue.ToString().Length;
                return System.Convert.ToInt16(newNum.ToString().Substring(0, carryNum + 1)) - System.Convert.ToInt16(strValue.ToString().Substring(0, 1));
            }

            return 0;
        }
    }
}