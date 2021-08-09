using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject
{
    public class ResultModelBase<T> where T : class, new()
    {
        public List<T> DataList { get; set; }
        public bool Result{ get; set; }
        public string ErrorMessage { get; set; }
        public Exception Exception { get; set; }

    }
}
