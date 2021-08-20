using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject
{
    public class BaseResult
    {
        public bool Result { get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }

        public static implicit operator bool(BaseResult dualResult)
        {
            return dualResult.Result;
        }
    }

    public class ResultModelBase<T> : BaseResult where T : class, new() 
    {
        public List<T> DataList { get; set; }
    }
}
