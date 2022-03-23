using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject
{
    public class BaseResult
    {
        public bool Result { get; set; } = true;

        private string _ErrorMeassage;
        public string ErrorMessage
        {
            get
            {
                return _ErrorMeassage;
            }
            set
            {
                _ErrorMeassage = string.IsNullOrEmpty(value) ? string.Empty : value.Replace("\r\n", "<br />");
            }
        }
        public Exception Exception { get; set; }
        public object ResultObject { get; set; }

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
