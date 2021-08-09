using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    /*(QualityPass1Provider) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Admin; Date: 2021/07/30  </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/07/30  1.00    Admin        Create
    /// </history>
    public interface IQualityPass1Provider
    {
        IList<Quality_Pass1> Get(Quality_Pass1 Item);
    }
}
