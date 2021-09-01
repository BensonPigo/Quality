using DatabaseObject.ProductionDB;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    /*Grey Scale 基本檔(ScaleProvider) 詳細敘述如下*/
    /// <summary>
    /// Grey Scale 基本檔
    /// </summary>
    /// <info>Author: Admin; Date: 2021/08/30  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/30  1.00    Admin        Create
    /// </history>
    public interface IScaleProvider
    {
        IList<Scale> Get();
    }
}
