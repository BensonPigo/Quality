using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    /*系統參數檔(SystemProvider) 詳細敘述如下*/
    /// <summary>
    /// 系統參數檔
    /// </summary>
    /// <info>Author: Admin; Date: 2021/08/12  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/12  1.00    Admin        Create
    /// </history>
    public interface ISystemProvider
    {
        IList<DatabaseObject.ProductionDB.System> Get();
    }
}
