using System;

namespace ADOHelper.Utility.Interface
{
	public interface ISQLDataTransaction
	{
		void CloseConnection();

		void Commit();

		void OpenConnection();

		void RollBack();
	}
}